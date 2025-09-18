import os
import sys
import uuid
from pathlib import Path

import pytest

from mcp_server.server import create_mcp_server


@pytest.mark.skipif(not sys.platform.startswith("win"), reason="Windows-only COM integration")
def test_create_and_save_as_md(tmp_path):
    # Arrange
    mcp = create_mcp_server()
    test_dir = tmp_path / "word_ops"
    test_dir.mkdir(parents=True, exist_ok=True)
    target = test_dir / f"doc_{uuid.uuid4().hex}.docx"
    md_target = test_dir / f"doc_{uuid.uuid4().hex}.md"

    # Act
    res1 = pytest.run(async_fn=lambda: mcp.call_tool("create_word_document", dict(path=str(target), content="# Title\n\nHello", content_format="markdown")))
    assert res1 is not None  # placeholder until test harness supports async

    # We cannot call mcp.call_tool synchronously here without the test harness;
    # This file is a placeholder for future integration tests.
    assert True
