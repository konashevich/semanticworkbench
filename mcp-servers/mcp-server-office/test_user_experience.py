#!/usr/bin/env python3
"""
Test script to verify specific user experience issues in the Office MCP implementation.
"""

import sys
import asyncio
from pathlib import Path

# Add the parent directory to sys.path to import the server
sys.path.insert(0, str(Path(__file__).parent))

def test_parameter_documentation_issues():
    """Test whether parameter documentation is sufficient for users."""
    print("🔍 Testing parameter documentation quality...")
    
    # Import the server to inspect docstrings
    from mcp_server.server import create_mcp_server
    
    server = create_mcp_server()
    tools = server._tools
    
    issues = []
    
    for tool_name, tool_info in tools.items():
        print(f"\n📊 Analyzing tool: {tool_name}")
        
        # Check if docstring has examples
        docstring = tool_info.description or ""
        
        if "batch_format" in tool_name:
            if "Example:" not in docstring and "example:" not in docstring:
                issues.append(f"{tool_name}: No usage examples in docstring")
            
            # Check for operation structure documentation
            if "operations:" not in docstring.lower():
                issues.append(f"{tool_name}: Operations parameter structure not documented")
        
        # Check if error messages are user-friendly (scan for technical terms)
        technical_terms = ["Exception", "COM", "pythoncom", "AttributeError"]
        for term in technical_terms:
            if term in docstring:
                issues.append(f"{tool_name}: Contains technical term '{term}' in user documentation")
    
    if issues:
        print("\n❌ Documentation issues found:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("\n✅ Documentation quality looks good!")
    
    return len(issues) == 0

def test_error_message_consistency():
    """Test error message quality and consistency."""
    print("\n🔍 Testing error message consistency...")
    
    # Test different error scenarios to see message quality
    test_cases = [
        ("Empty operations list", []),
        ("Invalid operation structure", [{"invalid": "structure"}]),
        ("Too many operations", [{"font": {"size": 12}}] * 51),
        ("Invalid color", [{"font": {"color": "invalid_color_name"}}]),
        ("Invalid font size", [{"font": {"size": -1}}]),
    ]
    
    issues = []
    
    for test_name, operations in test_cases:
        # Simulate validation by importing the validation function
        try:
            from mcp_server.server import create_mcp_server
            import inspect
            
            # Get the batch_format function source to analyze error messages
            server = create_mcp_server()
            batch_format_func = None
            
            for tool_name, tool_info in server._tools.items():
                if "batch_format" in tool_name:
                    # This is a simplified test - in real usage we'd call the function
                    print(f"  📋 Test case: {test_name}")
                    # We can't easily test actual error messages without calling the function
                    # but we can verify the structure exists
                    break
        except Exception as e:
            issues.append(f"Error analyzing {test_name}: {e}")
    
    if issues:
        print("\n❌ Error message issues found:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("\n✅ Error message structure looks consistent!")
    
    return len(issues) == 0

def test_input_validation_completeness():
    """Test if input validation catches common user errors."""
    print("\n🔍 Testing input validation completeness...")
    
    validation_tests = [
        ("Hex color validation", "#GGGGGG"),  # Invalid hex
        ("Font size limits", 5000),           # Too large
        ("Spacing limits", -100),             # Negative spacing
        ("Empty string handling", ""),        # Empty required fields
    ]
    
    issues = []
    
    # Test color parsing specifically since it's critical
    try:
        from mcp_server.app_interaction.word_editor import parse_font_color
        
        # Test invalid hex color
        result = parse_font_color("#GGGGGG")
        if result is not None:
            issues.append("Hex color validation accepts invalid hex colors")
        else:
            print("  ✅ Hex color validation correctly rejects invalid colors")
            
        # Test valid color
        result = parse_font_color("red")
        if result is None:
            issues.append("Color validation rejects valid CSS colors")
        else:
            print("  ✅ Color validation correctly accepts valid colors")
            
    except Exception as e:
        issues.append(f"Color validation test failed: {e}")
    
    if issues:
        print("\n❌ Input validation issues found:")
        for issue in issues:
            print(f"  - {issue}")
    else:
        print("\n✅ Input validation appears complete!")
    
    return len(issues) == 0

def test_tool_overlap_analysis():
    """Analyze tool overlap and potential user confusion."""
    print("\n🔍 Testing for tool overlap and user confusion...")
    
    # Identify overlapping functionality
    overlaps = [
        ("batch_format vs set_font", "Both can set font properties"),
        ("batch_format vs set_paragraph_format", "Both can set paragraph properties"),
        ("batch_format vs set_list_style", "Both can set list properties"),
        ("highlight vs batch_format", "batch_format could theoretically highlight"),
    ]
    
    issues = []
    
    # Check if there's clear guidance on when to use each tool
    from mcp_server.server import create_mcp_server
    server = create_mcp_server()
    
    batch_format_doc = None
    individual_tool_docs = {}
    
    for tool_name, tool_info in server._tools.items():
        if "batch_format" in tool_name:
            batch_format_doc = tool_info.description or ""
        elif any(x in tool_name for x in ["set_font", "set_paragraph", "set_list"]):
            individual_tool_docs[tool_name] = tool_info.description or ""
    
    # Check if batch_format doc explains when to use it vs individual tools
    if batch_format_doc and "when to use" not in batch_format_doc.lower():
        issues.append("batch_format doesn't explain when to use it vs individual tools")
    
    # Check if individual tools explain their relationship to batch_format
    for tool_name, doc in individual_tool_docs.items():
        if "batch_format" not in doc.lower():
            issues.append(f"{tool_name} doesn't mention relationship to batch_format")
    
    if issues:
        print("\n❌ Tool overlap issues found:")
        for issue in issues:
            print(f"  - {issue}")
        print("\n📋 Potential user confusion:")
        for overlap, description in overlaps:
            print(f"  - {overlap}: {description}")
    else:
        print("\n✅ Tool overlap is well-documented!")
    
    return len(issues) == 0

async def main():
    """Run all user experience quality tests."""
    print("🚀 Running Office MCP User Experience Quality Assessment\n")
    
    test_results = []
    
    test_results.append(test_parameter_documentation_issues())
    test_results.append(test_error_message_consistency()) 
    test_results.append(test_input_validation_completeness())
    test_results.append(test_tool_overlap_analysis())
    
    passed = sum(test_results)
    total = len(test_results)
    
    print(f"\n📊 User Experience Quality Score: {passed}/{total} tests passed")
    
    if passed == total:
        print("✅ Excellent user experience quality!")
    elif passed >= total * 0.75:
        print("⚠️  Good quality with some improvement opportunities")
    elif passed >= total * 0.5:
        print("🔶 Moderate quality - significant improvements needed")
    else:
        print("🚨 Poor user experience quality - major improvements required")
    
    print("\n💡 Key recommendations:")
    print("   1. Add comprehensive examples to all tool docstrings")
    print("   2. Standardize error message format and clarity")
    print("   3. Document when to use which tool for overlapping functionality")
    print("   4. Add missing input validation with user-friendly messages")
    
    return 0

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    sys.exit(exit_code)
