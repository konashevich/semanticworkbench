# Copyright (c) Microsoft. All rights reserved.

import pytest
import sys
from unittest.mock import Mock, patch, MagicMock
import asyncio

# Import the server functions we want to test
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from mcp_server.server import batch_format


class TestEnhancedBatchFormat:
    """Test suite for enhanced batch_format tool with atomic behavior and validation."""

    @pytest.mark.asyncio
    async def test_operation_limit_validation(self):
        """Test that batch_format rejects more than 50 operations."""
        # Create 51 operations
        operations = [{"font": {"bold": True}}] * 51
        
        result = await batch_format(
            scope="document",
            operations=operations
        )
        
        assert "error" in result
        assert "Maximum 50 operations allowed" in result["error"]

    @pytest.mark.asyncio
    async def test_font_size_validation(self):
        """Test font size validation (1-1638 points)."""
        # Test invalid font size - too small
        result = await batch_format(
            scope="document", 
            operations=[{"font": {"size": 0}}]
        )
        assert "error" in result
        assert "Font size must be between 1 and 1638" in result["error"]
        
        # Test invalid font size - too large
        result = await batch_format(
            scope="document",
            operations=[{"font": {"size": 2000}}]
        )
        assert "error" in result
        assert "Font size must be between 1 and 1638" in result["error"]

    @pytest.mark.asyncio
    async def test_paragraph_spacing_validation(self):
        """Test paragraph spacing validation (0-1584 points)."""
        # Test invalid space_before
        result = await batch_format(
            scope="document",
            operations=[{"paragraph": {"space_before": -1}}]
        )
        assert "error" in result
        assert "space_before must be between 0 and 1584" in result["error"]
        
        # Test invalid space_after
        result = await batch_format(
            scope="document",
            operations=[{"paragraph": {"space_after": 2000}}]
        )
        assert "error" in result
        assert "space_after must be between 0 and 1584" in result["error"]

    @pytest.mark.asyncio
    async def test_alignment_validation(self):
        """Test paragraph alignment validation."""
        result = await batch_format(
            scope="document",
            operations=[{"paragraph": {"alignment": "invalid_alignment"}}]
        )
        assert "error" in result
        assert "Invalid alignment" in result["error"]
        assert "left, center, right, justify" in result["error"]

    @pytest.mark.asyncio
    async def test_list_type_validation(self):
        """Test list type validation."""
        result = await batch_format(
            scope="document",
            operations=[{"list": {"type": "invalid_type"}}]
        )
        assert "error" in result
        assert "Invalid list type" in result["error"]
        assert "bullet, numbered, none" in result["error"]

    @pytest.mark.asyncio
    async def test_superscript_subscript_mutual_exclusion(self):
        """Test that superscript and subscript are mutually exclusive."""
        # Test within single operation
        result = await batch_format(
            scope="document",
            operations=[{"font": {"superscript": True, "subscript": True}}]
        )
        assert "error" in result
        assert "mutually exclusive" in result["error"]
        
        # Test across multiple operations
        result = await batch_format(
            scope="document",
            operations=[
                {"font": {"superscript": True}},
                {"font": {"subscript": True}}
            ]
        )
        assert "error" in result
        assert "mutually exclusive" in result["error"]

    @pytest.mark.asyncio
    async def test_performance_limit_max_matches(self):
        """Test performance limit on max_matches."""
        result = await batch_format(
            scope="matches",
            find={"text": "test", "max_matches": 15000},
            operations=[{"font": {"bold": True}}]
        )
        assert "error" in result
        assert "Maximum 10,000 matches allowed" in result["error"]

    @pytest.mark.asyncio
    async def test_matches_scope_requires_find_text(self):
        """Test that matches scope requires find text."""
        result = await batch_format(
            scope="matches",
            operations=[{"font": {"bold": True}}]
        )
        assert "error" in result
        assert "provide find.text" in result["error"]

    @pytest.mark.asyncio
    async def test_invalid_scope(self):
        """Test invalid scope validation."""
        result = await batch_format(
            scope="invalid_scope",
            operations=[{"font": {"bold": True}}]
        )
        assert "error" in result
        assert "Invalid scope" in result["error"]
        assert "document, selection, or matches" in result["error"]

    @pytest.mark.asyncio
    async def test_clear_formatting_target_validation(self):
        """Test clear_formatting target validation."""
        result = await batch_format(
            scope="document",
            operations=[{"clear_formatting": {"target": "invalid_target"}}]
        )
        assert "error" in result
        assert "Invalid clear_formatting target" in result["error"]
        assert "all, font, paragraph" in result["error"]

    @pytest.mark.asyncio
    async def test_document_base_font_validation(self):
        """Test document base font validation."""
        # Test invalid size
        result = await batch_format(
            scope="document",
            operations=[{"document_base_font": {"size": 0}}]
        )
        assert "error" in result
        assert "Document base font size must be between 1 and 1638" in result["error"]

    @pytest.mark.asyncio
    async def test_empty_operations_list(self):
        """Test behavior with empty operations list."""
        result = await batch_format(
            scope="document",
            operations=[]
        )
        assert result["changed"] is False
        assert result["summary"]["operations"] == 0
        assert result["summary"]["ranges_affected"] == 0
        assert result["details"] == []

    @pytest.mark.asyncio
    async def test_none_operations(self):
        """Test behavior with None operations."""
        result = await batch_format(
            scope="document",
            operations=None
        )
        assert result["changed"] is False
        assert result["summary"]["operations"] == 0

    @pytest.mark.asyncio  
    async def test_find_dict_format(self):
        """Test that find parameter accepts dictionary format."""
        # This test would require mocking Word COM objects
        # For now, just test that the validation passes
        result = await batch_format(
            scope="matches",
            find={"text": "test", "case_sensitive": True, "whole_word": False},
            operations=[{"font": {"bold": True}}]
        )
        # This will fail because we don't have Word available in test,
        # but it should pass validation
        assert "error" in result
        # Should not be a validation error about missing find.text
        assert "provide find.text" not in result["error"]


class TestEnhancedHighlight:
    """Test suite for enhanced highlight tool."""

    @pytest.mark.asyncio
    async def test_invalid_scope(self):
        """Test invalid scope validation."""
        from mcp_server.server import highlight
        
        result = await highlight(
            text="test",
            scope="invalid_scope"
        )
        assert "error" in result
        assert "Invalid scope" in result["error"]
        assert "first or all" in result["error"]

    @pytest.mark.asyncio 
    async def test_empty_text_validation(self):
        """Test empty text validation."""
        from mcp_server.server import highlight
        
        result = await highlight(text="")
        assert "error" in result
        assert "cannot be empty" in result["error"]

    @pytest.mark.asyncio
    async def test_performance_limit_max_matches(self):
        """Test performance limit validation."""
        from mcp_server.server import highlight
        
        result = await highlight(
            text="test",
            max_matches=15000
        )
        assert "error" in result
        assert "Maximum 10,000 matches allowed" in result["error"]

    @pytest.mark.asyncio
    async def test_invalid_color_validation(self):
        """Test invalid color validation."""
        from mcp_server.server import highlight
        
        result = await highlight(
            text="test",
            color="invalid_color_name"
        )
        assert "error" in result
        assert "Invalid highlight color" in result["error"]


if __name__ == "__main__":
    # Run tests for manual verification
    pytest.main([__file__, "-v"])
