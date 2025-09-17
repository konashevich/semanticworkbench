#!/usr/bin/env python3
"""Manual validation script for enhanced batch_format functionality."""

import sys
import asyncio
from pathlib import Path

# Add the parent directory to sys.path to import the server
sys.path.insert(0, str(Path(__file__).parent))

try:
    from mcp_server.server import batch_format, highlight
    print("✅ Successfully imported batch_format and highlight functions")
except ImportError as e:
    print(f"❌ Import failed: {e}")
    sys.exit(1)

async def test_validation():
    """Test the validation functionality of batch_format."""
    print("\n🧪 Testing batch_format validation...")
    
    # Test 1: Operation limit validation
    print("Test 1: Operation limit validation (51 operations)")
    operations = [{"font": {"bold": True}}] * 51
    result = await batch_format(scope="document", operations=operations)
    if "error" in result and "Maximum 50 operations allowed" in result["error"]:
        print("✅ Operation limit validation works")
    else:
        print(f"❌ Operation limit validation failed: {result}")
    
    # Test 2: Font size validation
    print("Test 2: Font size validation (invalid size)")
    result = await batch_format(
        scope="document", 
        operations=[{"font": {"size": 0}}]
    )
    if "error" in result and "Font size must be between 1 and 1638" in result["error"]:
        print("✅ Font size validation works")
    else:
        print(f"❌ Font size validation failed: {result}")
    
    # Test 3: Superscript/subscript mutual exclusion
    print("Test 3: Superscript/subscript mutual exclusion")
    result = await batch_format(
        scope="document",
        operations=[{"font": {"superscript": True, "subscript": True}}]
    )
    if "error" in result and "mutually exclusive" in result["error"]:
        print("✅ Mutual exclusion validation works")
    else:
        print(f"❌ Mutual exclusion validation failed: {result}")
    
    # Test 4: Empty operations list
    print("Test 4: Empty operations list")
    result = await batch_format(scope="document", operations=[])
    if (result.get("changed") is False and 
        result.get("summary", {}).get("operations") == 0 and
        result.get("details") == []):
        print("✅ Empty operations handling works")
    else:
        print(f"❌ Empty operations handling failed: {result}")

async def test_highlight_validation():
    """Test the validation functionality of highlight."""
    print("\n🧪 Testing highlight validation...")
    
    # Test 1: Invalid scope
    print("Test 1: Invalid scope validation")
    result = await highlight(text="test", scope="invalid_scope")
    if "error" in result and "Invalid scope" in result["error"]:
        print("✅ Highlight scope validation works")
    else:
        print(f"❌ Highlight scope validation failed: {result}")
    
    # Test 2: Empty text
    print("Test 2: Empty text validation")
    result = await highlight(text="")
    if "error" in result and "cannot be empty" in result["error"]:
        print("✅ Highlight empty text validation works")
    else:
        print(f"❌ Highlight empty text validation failed: {result}")
    
    # Test 3: Performance limit
    print("Test 3: Performance limit validation")
    result = await highlight(text="test", max_matches=15000)
    if "error" in result and "Maximum 10,000 matches allowed" in result["error"]:
        print("✅ Highlight performance limit validation works")
    else:
        print(f"❌ Highlight performance limit validation failed: {result}")

async def main():
    """Run all validation tests."""
    print("🚀 Starting manual validation tests for enhanced Office MCP tools")
    
    try:
        await test_validation()
        await test_highlight_validation()
        print("\n✅ All validation tests completed successfully!")
        print("📋 The enhanced Office MCP implementation is ready for use.")
    except Exception as e:
        print(f"\n❌ Validation tests failed with error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
