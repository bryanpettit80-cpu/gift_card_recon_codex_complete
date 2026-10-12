"""Single-workstation writer coordination and crash-recoverable publication.

An incomplete transaction restores its recorded preimages. A committed transaction
finishes fingerprint-checked cleanup. Neither recovery path needs the original
staging directory or live POS source. Dropbox handoff requires completed sync.
"""
from __future__ import annotations

import errno
import hashlib
import json
import os
import shutil
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable


class RecoveryRequired(RuntimeError):
    pass


_locks: dict[str, threading.RLock] = {}
_depth = threading.local()
_active: set[Path] = set()
_UNSPECIFIED = object()


def support_root_for(archive_root: Path) -> Path:
    archive_root = Path(archive_root).resolve()
    for candidate in (archive_root, *archive_root.parents):
        if candidate.name.casefold() in {"04 archive", "archive - old files"}:
            return candidate.parent / "_automation_runs" / "support"
    return archive_root.parent / "_automation_runs" / "support"


@contextmanager
def writer_lock(support_root: Path):
    root = Path(support_root).resolve()
    key = str(root).casefold() if os.name == "nt" else str(root)
    gate = _locks.setdefault(key, threading.RLock())
    if not gate.acquire(blocking=False):
        raise RuntimeError("Another gift-card reporting writer is active in this workspace.")
    depths = getattr(_depth, "values", {})
    _depth.values = depths
    stream = None
    locked = False
    entered = False
    try:
        if depths.get(key, 0) == 0:
            root.mkdir(parents=True, exist_ok=True)
            stream = (root / "writer.lock").open("a+b")
            stream.seek(0, os.SEEK_END)
            if stream.tell() == 0:
                stream.write(b"0")
                stream.flush()
            stream.seek(0)
            try:
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                locked = True
            except OSError as exc:
                raise RuntimeError("Another gift-card reporting writer is active in this workspace.") from exc
        depths[key] = depths.get(key, 0) + 1
        entered = True
        yield
    finally:
        if entered:
            depths[key] -= 1
        if stream is not None:
            if locked:
                stream.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
            stream.close()
        gate.release()


def retry_file_operation(operation):
    delays = (0.10, 0.25, 0.50, 1.00, 2.00)
    for attempt in range(len(delays) + 1):
        try:
            return operation()
        except OSError as exc:
            winerror = getattr(exc, "winerror", None)
            transient = winerror in {5, 32, 33, 148, 170} if winerror is not None else exc.errno in {errno.EACCES, errno.EBUSY}
            if not transient or attempt == len(delays):
                raise
            time.sleep(delays[attempt])


def _digest(path: Path):
    if path.is_symlink():
        raise RecoveryRequired(f"Publication does not follow symbolic links: {path}")
    if not path.exists():
        return None
    digest = hashlib.sha256()
    if path.is_dir():
        for item in sorted(path.rglob("*")):
            if item.is_symlink():
                raise RecoveryRequired(f"Publication does not follow symbolic links: {item}")
            digest.update(item.relative_to(path).as_posix().encode("utf-8"))
            digest.update(b"D" if item.is_dir() else b"F")
            if item.is_file():
                digest.update(_digest(item).encode("ascii"))
        return "dir:" + digest.hexdigest()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return "file:" + digest.hexdigest()


def file_state(path: Path) -> str | None:
    """Capture an expected publication preimage, including an absent destination."""
    return _digest(Path(path))


def _copy_file_durable(source, destination):
    shutil.copyfile(source, destination)
    with Path(destination).open("r+b") as stream:
        os.fsync(stream.fileno())
    # Apply read-only attributes only after flushing the new writable copy.
    shutil.copystat(source, destination)
    return destination


def _copy(source: Path, destination: Path):
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir():
        shutil.copytree(source, destination, copy_function=_copy_file_durable)
    else:
        _copy_file_durable(source, destination)
    if _digest(source) != _digest(destination):
        raise RecoveryRequired(f"Publication copy did not verify: {destination}")


def _remove(path: Path):
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink(missing_ok=True)


def _write(path: Path, payload):
    temporary = path.with_suffix(".tmp")
    with temporary.open("w", encoding="utf-8") as stream:
        json.dump(payload, stream, sort_keys=True, indent=2)
        stream.flush()
        os.fsync(stream.fileno())
    retry_file_operation(lambda: os.replace(temporary, path))


def _checkpoint(phase: str):
    """Boundary hook used by subprocess termination tests."""


class PublicationTransaction:
    def __init__(self, support_root: Path):
        support_root = Path(support_root).resolve()
        self.root = support_root / "transactions" / uuid.uuid4().hex
        self.root.mkdir(parents=True)
        self.path = self.root / "transaction.json"
        self.data = {"schema_version": 1, "id": self.root.name, "state": "active", "operations": [], "cleanup": [],
                     "support_root": str(support_root), "workspace_root": str(_workspace_anchor(support_root))}
        _write(self.path, self.data)

    def __enter__(self):
        _active.add(self.path)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.data["state"] == "active":
                self.rollback()
        finally:
            _active.discard(self.path)

    @property
    def committed(self):
        return self.data["state"] in {"committed", "completed"}

    def _operation(self, source: Path | None, destination: Path, expected_before=_UNSPECIFIED):
        if self.data["state"] != "active":
            raise RuntimeError("Cannot modify a completed publication transaction.")
        destination = Path(destination).absolute()
        if any(item.is_symlink() for item in (destination, *destination.parents)):
            raise RecoveryRequired(f"Publication does not follow symbolic links: {destination}")
        destination = destination.resolve()
        if any(Path(op["destination"]) == destination for op in self.data["operations"]):
            raise ValueError(f"A publication destination may only be changed once: {destination}")
        before = _digest(destination)
        if expected_before is not _UNSPECIFIED and before != expected_before:
            raise RecoveryRequired(f"Destination changed since validation; retained for review: {destination}")
        after = _digest(source) if source is not None else None
        if source is not None and after is None:
            raise FileNotFoundError(source)
        backup = self.root / "backup" / str(len(self.data["operations"]))
        if before is not None:
            _copy(destination, backup)
            if _digest(backup) != before:
                raise RecoveryRequired(f"Destination changed while capturing previous result: {destination}")
        prefix = ".gcr-" + self.root.name[:12] + "-" + str(len(self.data["operations"]))
        staged = destination.parent / (prefix + ".stage")
        displaced = destination.parent / (prefix + ".previous")
        discarded = destination.parent / (prefix + ".discarded")
        created_parents = []
        parent = destination.parent
        while not parent.exists():
            created_parents.append(str(parent))
            parent = parent.parent
        operation = {"destination": str(destination), "backup": str(backup),
                     "staged": str(staged), "displaced": str(displaced),
                     "discarded": str(discarded), "created_parents": created_parents,
                     "before": before, "after": after}
        self.data["operations"].append(operation)
        _write(self.path, self.data)
        _checkpoint("intent")
        if _digest(destination) != before:
            raise RecoveryRequired(f"Destination changed before publication: {destination}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        if source is not None:
            _copy(source, staged)
            if _digest(staged) != after:
                raise RecoveryRequired(f"Source changed while preparing publication: {source}")
            _checkpoint("staged")
        if _digest(destination) != before:
            raise RecoveryRequired(f"Destination changed before replacement: {destination}")
        if destination.is_dir():
            retry_file_operation(lambda: os.replace(destination, displaced))
            _checkpoint("displaced")
        if source is not None:
            retry_file_operation(lambda: os.replace(staged, destination))
        else:
            retry_file_operation(lambda: _remove(destination))
        _checkpoint("replaced")
        if _digest(destination) != after:
            raise RecoveryRequired(f"Published artifact verification failed: {destination}")

    def replace(self, staged: Path, destination: Path, *, expected_before=_UNSPECIFIED):
        self._operation(Path(staged), Path(destination), expected_before)

    def remove(self, path: Path, *, expected_before=_UNSPECIFIED):
        self._operation(None, Path(path), expected_before)

    def commit(self, cleanup: Iterable[Path] = (), *, expected_cleanup=None) -> list[str]:
        expected = {str(Path(p).resolve()): value for p, value in (expected_cleanup or {}).items()}
        committed = dict(self.data)
        committed["cleanup"] = [
            {"path": str(Path(p).resolve()),
             "digest": "file:" + expected[str(Path(p).resolve())]
             if str(Path(p).resolve()) in expected else _digest(Path(p))}
            for p in cleanup
        ]
        committed["state"] = "committed"
        _write(self.path, committed)
        self.data = committed
        _checkpoint("committed")
        try:
            return _finish_cleanup(self.path, self.data)
        except (OSError, RecoveryRequired) as exc:
            # The result is already durable. A failed cleanup/state update must
            # not turn it into a reported rollback or discard its recovery record.
            return [f"Publication committed; cleanup recovery pending at {self.path}: {exc}"]

    def rollback(self):
        _rollback(self.path, self.data)


def _restore_operation(path: Path, operation):
    destination = Path(operation["destination"])
    before, after = operation["before"], operation["after"]
    current = _digest(destination)
    if current == before:
        return
    if current is not None and current != after:
        raise RecoveryRequired(f"Recovery retained a changed destination for review: {destination}. Journal: {path}")
    discarded = Path(operation["discarded"])
    if before is None:
        if destination.is_dir():
            retry_file_operation(lambda: os.replace(destination, discarded))
        else:
            retry_file_operation(lambda: _remove(destination))
    else:
        backup = Path(operation["backup"])
        if _digest(backup) != before:
            raise RecoveryRequired(f"Previous-good publication backup failed verification: {backup}")
        restored = Path(operation["staged"])
        # A prior attempt may have left a partial staging copy. This is a recorded,
        # application-owned path and never an operator input or published result.
        retry_file_operation(lambda: _remove(restored))
        _copy(backup, restored)
        if destination.is_dir():
            retry_file_operation(lambda: os.replace(destination, discarded))
        destination.parent.mkdir(parents=True, exist_ok=True)
        retry_file_operation(lambda: os.replace(restored, destination))
    if _digest(destination) != before:
        raise RecoveryRequired(f"Publication restore failed: {destination}")
    _checkpoint("restored")


def _rollback(path: Path, data):
    errors = []
    for operation in reversed(data["operations"]):
        try:
            _restore_operation(path, operation)
        except (OSError, RuntimeError, ValueError) as exc:
            errors.append(f"retained path {operation['destination']}: {exc}")
    if errors:
        raise RecoveryRequired("Rollback incomplete: " + "; ".join(errors) + f". Recovery journal: {path}")
    data["state"] = "rolled_back"
    data["housekeeping_pending"] = True
    _write(path, data)
    _discard_backups(path, data)


def _discard_backups(path, data):
    warnings = []
    targets = [path.parent / "backup"] + [Path(operation[name])
        for operation in data["operations"] for name in ("staged", "displaced", "discarded")]
    for target in targets:
        try:
            retry_file_operation(lambda target=target: _remove(target))
        except OSError as exc:
            warnings.append(f"Publication housekeeping pending at {target}: {exc}")
    if data["state"] == "rolled_back":
        parents = {Path(parent) for op in data["operations"] for parent in op.get("created_parents", [])}
        for parent in sorted(parents, key=lambda p: len(p.parts), reverse=True):
            try:
                if parent.is_dir() and not any(parent.iterdir()):
                    parent.rmdir()
            except OSError as exc:
                warnings.append(f"Publication housekeeping pending at {parent}: {exc}")
    data["housekeeping_pending"] = bool(warnings)
    _write(path, data)
    return warnings


def _finish_cleanup(path: Path, data) -> list[str]:
    warnings = []
    for entry in data["cleanup"]:
        target = Path(entry["path"])
        if entry.get("unmapped_after_move"):
            warnings.append(f"Cleanup pending for external source after workspace move: {target}")
            continue
        try:
            current = _digest(target)
            if current is None:
                continue
            if current != entry["digest"]:
                warnings.append(f"Cleanup pending; changed input retained: {target}")
                continue
            def remove_unchanged():
                if _digest(target) not in (None, entry["digest"]):
                    raise RecoveryRequired(f"Cleanup input changed: {target}")
                _remove(target)
            retry_file_operation(remove_unchanged)
        except (OSError, RecoveryRequired) as exc:
            warnings.append(f"Cleanup pending at {target}: {exc}")
    if not warnings:
        data["state"] = "completed"
        data["housekeeping_pending"] = True
        _write(path, data)
        warnings.extend(_discard_backups(path, data))
    return warnings


def recover_pending(support_root: Path) -> list[str]:
    actions = []
    for path in sorted((Path(support_root).resolve() / "transactions").glob("*/transaction.json")):
        if path in _active:
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("schema_version") != 1:
            raise RecoveryRequired(f"Unsupported recovery journal: {path}")
        if data["state"] in {"active", "committed"} or data.get("housekeeping_pending"):
            _rebase_journal(path, data, Path(support_root).resolve())
        if data["state"] == "active":
            _rollback(path, data)
            actions.append(f"Restored interrupted publication {data['id']}.")
        elif data["state"] == "committed":
            warnings = _finish_cleanup(path, data)
            actions.extend(warnings or [f"Finished committed publication cleanup {data['id']}."])
        elif data.get("housekeeping_pending"):
            warnings = _discard_backups(path, data)
            actions.extend(warnings or [f"Finished publication housekeeping {data['id']}."])
    return actions


def _workspace_anchor(support: Path) -> Path:
    return support.parent.parent if support.parent.name == "_automation_runs" else support.parent


def _rebase_journal(path, data, support):
    old_support = Path(data.get("support_root", support))
    if old_support == support:
        return
    previous = Path(data["workspace_root"])
    current = _workspace_anchor(support)
    def mapped(value):
        candidate = Path(value)
        try:
            return str(current / candidate.relative_to(previous))
        except ValueError:
            raise RecoveryRequired("Interrupted publication uses an external path that cannot be remapped after a move: "
                                   f"{candidate}. Recover with the original configured roots first.")
    for operation in data["operations"]:
        for key in ("destination", "backup", "staged", "displaced", "discarded"):
            operation[key] = mapped(operation[key])
        operation["created_parents"] = [mapped(value) for value in operation.get("created_parents", [])]
    for entry in data["cleanup"]:
        try:
            entry["path"] = mapped(entry["path"])
        except RecoveryRequired:
            entry["unmapped_after_move"] = True
    data["support_root"], data["workspace_root"] = str(support), str(current)
    _write(path, data)
