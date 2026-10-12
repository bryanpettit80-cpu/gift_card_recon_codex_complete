from __future__ import annotations

import errno
import json
import os
import subprocess
import sys
import shutil
from pathlib import Path

import pytest

from gift_card_recon import publication
from gift_card_recon.publication import PublicationTransaction, recover_pending, writer_lock
from gift_card_recon.utils import sha256_file


@pytest.mark.parametrize("boundary", ["intent", "staged", "displaced", "replaced", "committed"])
def test_process_death_restores_previous_or_finishes_committed_cleanup(tmp_path, boundary):
    old = tmp_path / "report"
    old.mkdir()
    (old / "workbook").write_text("previous-good")
    new = tmp_path / "new"
    new.mkdir()
    (new / "workbook").write_text("corrected")
    inbox = tmp_path / "input"
    inbox.write_text("source")
    script = """
import os,sys
from pathlib import Path
from gift_card_recon import publication as p
root=Path(sys.argv[1]); boundary=sys.argv[2]
def checkpoint(phase):
    if phase == boundary: os._exit(79)
p._checkpoint=checkpoint
with p.writer_lock(root/'support'):
    with p.PublicationTransaction(root/'support') as tx:
        tx.replace(root/'new', root/'report')
        tx.commit(cleanup=(root/'input',))
"""
    process = subprocess.run([sys.executable, "-c", script, str(tmp_path), boundary], timeout=30)
    assert process.returncode == 79
    with writer_lock(tmp_path / "support"):
        actions = recover_pending(tmp_path / "support")
        assert actions
        assert recover_pending(tmp_path / "support") == []
    assert (old / "workbook").read_text() == ("corrected" if boundary == "committed" else "previous-good")
    assert inbox.exists() == (boundary != "committed")
    assert not list(tmp_path.glob(".gcr-*"))


def test_locked_rollback_retains_journal_and_restores_on_next_run(tmp_path, monkeypatch):
    original = tmp_path / "report.txt"
    original.write_text("previous")
    staged = tmp_path / "stage.txt"
    staged.write_text("new")
    tx = PublicationTransaction(tmp_path / "support")
    tx.replace(staged, original)
    real_replace = publication.os.replace
    def locked(source, destination):
        if Path(destination) == original:
            raise OSError(errno.EBUSY, "locked")
        return real_replace(source, destination)
    with monkeypatch.context() as patch:
        patch.setattr(publication.os, "replace", locked)
        patch.setattr(publication.time, "sleep", lambda _: None)
        with pytest.raises(publication.RecoveryRequired, match="Rollback incomplete"):
            tx.rollback()
    assert original.read_text() == "new"
    with writer_lock(tmp_path / "support"):
        recover_pending(tmp_path / "support")
    assert original.read_text() == "previous"


def test_cross_volume_sources_are_copied_before_same_directory_publication(tmp_path, monkeypatch):
    source = tmp_path / "different-volume" / "source"
    source.parent.mkdir()
    source.write_text("new")
    destination = tmp_path / "output" / "report"
    destination.parent.mkdir()
    destination.write_text("old")
    real_replace = publication.os.replace
    def simulated_cross_device(source_path, destination_path):
        if Path(source_path).parent != Path(destination_path).parent:
            raise OSError(errno.EXDEV, "simulated cross-volume move")
        return real_replace(source_path, destination_path)
    monkeypatch.setattr(publication.os, "replace", simulated_cross_device)
    with writer_lock(tmp_path / "support"):
        with PublicationTransaction(tmp_path / "support") as tx:
            tx.replace(source, destination)
            tx.commit()
    assert destination.read_text() == "new"
    assert source.read_text() == "new"


def test_cleanup_never_removes_changed_source_at_commit_or_recovery(tmp_path):
    inbox = tmp_path / "inbox"
    inbox.write_text("calculated source")
    expected = sha256_file(inbox)
    inbox.write_text("new delivery")
    with writer_lock(tmp_path / "support"):
        with PublicationTransaction(tmp_path / "support") as tx:
            warnings = tx.commit(cleanup=(inbox,), expected_cleanup={inbox: expected})
        assert warnings
        assert recover_pending(tmp_path / "support")
    assert inbox.read_text() == "new delivery"


def test_writer_lock_rejects_other_process_and_is_released_after_termination(tmp_path):
    child = """
import sys
from pathlib import Path
from gift_card_recon.publication import writer_lock
try:
    with writer_lock(Path(sys.argv[1])): pass
except RuntimeError:
    sys.exit(42)
"""
    with writer_lock(tmp_path):
        with writer_lock(tmp_path):  # Same-thread service nesting is safe.
            result = subprocess.run([sys.executable, "-c", child, str(tmp_path)], timeout=30)
            assert result.returncode == 42
    assert subprocess.run([sys.executable, "-c", child, str(tmp_path)], timeout=30).returncode == 0


def test_unknown_destination_edit_is_retained_during_recovery(tmp_path):
    staged, output = tmp_path / "stage", tmp_path / "output"
    staged.write_text("new")
    output.write_text("old")
    tx = PublicationTransaction(tmp_path / "support")
    tx.replace(staged, output)
    output.write_text("operator edit after interruption")
    with pytest.raises(publication.RecoveryRequired, match="changed destination"):
        recover_pending(tmp_path / "support")
    assert output.read_text() == "operator edit after interruption"
    assert json.loads(tx.path.read_text())["state"] == "active"


def test_interrupted_workspace_copy_recovers_only_new_roots(tmp_path):
    old = tmp_path / "old"
    old.mkdir()
    (old / "output").write_text("previous")
    (old / "stage").write_text("replacement")
    tx = PublicationTransaction(old / "support")
    tx.replace(old / "stage", old / "output")
    new = tmp_path / "new-user"
    shutil.copytree(old, new)
    with writer_lock(new / "support"):
        recover_pending(new / "support")
    assert (new / "output").read_text() == "previous"
    assert (old / "output").read_text() == "replacement"


def test_completed_publication_retries_locked_backup_housekeeping(tmp_path, monkeypatch):
    output, source = tmp_path / "output", tmp_path / "source"
    output.write_text("previous")
    source.write_text("new")
    real_remove = publication._remove
    def locked_backup(path):
        if Path(path).name == "backup":
            raise PermissionError(errno.EACCES, "locked backup")
        return real_remove(path)
    with writer_lock(tmp_path / "support"):
        with PublicationTransaction(tmp_path / "support") as tx:
            tx.replace(source, output)
            with monkeypatch.context() as patch:
                patch.setattr(publication, "_remove", locked_backup)
                patch.setattr(publication.time, "sleep", lambda _: None)
                assert tx.commit()
        assert tx.committed
        assert json.loads(tx.path.read_text())["housekeeping_pending"]
        assert output.read_text() == "new"
        assert recover_pending(tmp_path / "support")
        assert not (tx.root / "backup").exists()
        assert not json.loads(tx.path.read_text())["housekeeping_pending"]
        assert recover_pending(tmp_path / "support") == []


def test_committed_result_remains_successful_when_cleanup_journal_write_fails(tmp_path, monkeypatch):
    output, source = tmp_path / "output", tmp_path / "source"
    output.write_text("previous")
    source.write_text("new")
    real_write = publication._write
    def unavailable(path, data):
        if data["state"] == "completed":
            raise OSError(errno.ENOSPC, "storage unavailable after commit")
        real_write(path, data)
    with writer_lock(tmp_path / "support"):
        with PublicationTransaction(tmp_path / "support") as tx:
            tx.replace(source, output)
            with monkeypatch.context() as patch:
                patch.setattr(publication, "_write", unavailable)
                warnings = tx.commit()
        assert tx.committed and "committed" in warnings[0]
        assert output.read_text() == "new"
        assert json.loads(tx.path.read_text())["state"] == "committed"
        assert recover_pending(tmp_path / "support")
        assert output.read_text() == "new"
