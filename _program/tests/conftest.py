"""Keep synthetic Windows operator paths within ordinary filesystem limits."""
from __future__ import annotations

import sys

import pytest


if sys.platform == "win32":
    @pytest.fixture
    def tmp_path(tmp_path_factory):
        # Fixture paths already contain realistic user/Dropbox/archive names.
        # Pytest's full test-name directory adds unrelated length and can make
        # CopyFile2 fail before the behavior under test is reached.
        yield tmp_path_factory.mktemp("t")
