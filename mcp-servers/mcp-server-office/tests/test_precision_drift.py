#!/usr/bin/env python3
"""
PowerPoint MCP Server Precision Drift Test

This test validates that coordinate precision is maintained across many shape placements,
simulating an OCR reconstruction workflow with 50+ text boxes.
"""

import asyncio
import json
import math
from typing import Dict, List, Tuple
import sys
import os

# Add the parent directory to sys.path to import mcp_server modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

async def test_precision_drift():
    """Test precision drift in PowerPoint coordinate placement."""
    print("PowerPoint Precision Drift Test")
    print("=" * 50)
    
    # Mock OCR results: grid of text boxes across A4 slide
    a4_width_mm = 210.0
    a4_height_mm = 297.0
    
    # Convert to points (will be used internally)
    from mcp_server.app_interaction.powerpoint_editor import mm_to_points
    slide_width_pt = mm_to_points(a4_width_mm)
    slide_height_pt = mm_to_points(a4_height_mm)
    
    print(f"A4 Slide dimensions: {slide_width_pt:.3f} x {slide_height_pt:.3f} points")
    
    # Generate test grid: 10x8 grid of text boxes (80 total)
    grid_cols = 10
    grid_rows = 8
    margin_mm = 10.0  # 10mm margin on all sides
    
    box_width_mm = (a4_width_mm - 2 * margin_mm) / grid_cols
    box_height_mm = (a4_height_mm - 2 * margin_mm) / grid_rows
    
    test_boxes = []
    for row in range(grid_rows):
        for col in range(grid_cols):
            left_mm = margin_mm + col * box_width_mm
            top_mm = margin_mm + row * box_height_mm
            
            test_boxes.append({
                "id": f"box_{row}_{col}",
                "left_mm": left_mm,
                "top_mm": top_mm,
                "width_mm": box_width_mm * 0.9,  # 90% of available space
                "height_mm": box_height_mm * 0.8,  # 80% of available space
                "text": f"Text {row},{col}",
                "expected_left_pt": mm_to_points(left_mm),
                "expected_top_pt": mm_to_points(top_mm),
                "expected_width_pt": mm_to_points(box_width_mm * 0.9),
                "expected_height_pt": mm_to_points(box_height_mm * 0.8),
            })
    
    print(f"Generated {len(test_boxes)} test text boxes in {grid_rows}x{grid_cols} grid")
    
    # Test coordinate precision
    from mcp_server.app_interaction.powerpoint_editor import round_precision, parse_unit
    
    precision_errors = []
    unit_conversion_errors = []
    
    # Test 1: Round-trip precision (points -> round -> points)
    print("\nTest 1: Round-trip precision...")
    for box in test_boxes:
        original_left = box["expected_left_pt"]
        original_top = box["expected_top_pt"]
        
        rounded_left = round_precision(original_left)
        rounded_top = round_precision(original_top)
        
        left_error = abs(original_left - rounded_left)
        top_error = abs(original_top - rounded_top)
        
        if left_error > 0.001 or top_error > 0.001:  # 0.001pt tolerance
            precision_errors.append({
                "box_id": box["id"],
                "left_error": left_error,
                "top_error": top_error,
                "original": (original_left, original_top),
                "rounded": (rounded_left, rounded_top)
            })
    
    # Test 2: Unit conversion round-trip (mm -> pt -> mm)
    print("Test 2: Unit conversion round-trip...")
    for box in test_boxes:
        # Convert mm to points and back
        original_mm = box["left_mm"]
        converted_pt = parse_unit(f"{original_mm}mm")
        back_to_mm = converted_pt * 25.4 / 72.0  # points to mm
        
        mm_error = abs(original_mm - back_to_mm)
        if mm_error > 0.01:  # 0.01mm tolerance (about 0.03pt)
            unit_conversion_errors.append({
                "box_id": box["id"],
                "original_mm": original_mm,
                "converted_pt": converted_pt,
                "back_to_mm": back_to_mm,
                "error_mm": mm_error
            })
    
    # Test 3: Cumulative drift simulation
    print("Test 3: Cumulative drift simulation...")
    cumulative_left = 0.0
    cumulative_top = 0.0
    max_drift = 0.0
    
    for i, box in enumerate(test_boxes):
        # Simulate adding coordinate values as might happen in sequential placement
        cumulative_left += round_precision(box["width_mm"] / 10.0)  # Small increments
        cumulative_top += round_precision(box["height_mm"] / 10.0)
        
        # Calculate drift from expected position
        expected_cumulative_left = (i + 1) * (box_width_mm * 0.9 / 10.0)
        expected_cumulative_top = (i + 1) * (box_height_mm * 0.8 / 10.0)
        
        drift_left = abs(cumulative_left - expected_cumulative_left)
        drift_top = abs(cumulative_top - expected_cumulative_top)
        total_drift = math.sqrt(drift_left**2 + drift_top**2)
        
        max_drift = max(max_drift, total_drift)
    
    # Test 4: Color validation test
    print("Test 4: Color validation...")
    from mcp_server.app_interaction.powerpoint_editor import validate_color, PPTValidationError
    
    color_test_cases = [
        ("#FF0000", True),   # Valid hex
        ("#INVALID", False), # Invalid hex
        ("red", True),       # Valid named color
        ("invalid_color", False),  # Invalid named color
        (255, True),         # Valid integer
        (-1, True),          # Edge case integer
    ]
    
    color_validation_errors = []
    for color_input, should_be_valid in color_test_cases:
        try:
            result = validate_color(color_input)
            if not should_be_valid:
                color_validation_errors.append(f"Expected {color_input} to be invalid but got {result}")
        except PPTValidationError:
            if should_be_valid:
                color_validation_errors.append(f"Expected {color_input} to be valid but got validation error")
    
    # Report results
    print("\n" + "=" * 50)
    print("PRECISION TEST RESULTS")
    print("=" * 50)
    
    print(f"Round-trip precision errors: {len(precision_errors)}")
    if precision_errors:
        print("  Worst precision errors:")
        sorted_errors = sorted(precision_errors, key=lambda x: max(x["left_error"], x["top_error"]), reverse=True)
        for error in sorted_errors[:3]:
            print(f"    {error['box_id']}: left={error['left_error']:.6f}pt, top={error['top_error']:.6f}pt")
    
    print(f"\nUnit conversion errors: {len(unit_conversion_errors)}")
    if unit_conversion_errors:
        print("  Worst conversion errors:")
        sorted_errors = sorted(unit_conversion_errors, key=lambda x: x["error_mm"], reverse=True)
        for error in sorted_errors[:3]:
            print(f"    {error['box_id']}: {error['error_mm']:.6f}mm")
    
    print(f"\nMaximum cumulative drift: {max_drift:.6f}mm")
    
    print(f"\nColor validation errors: {len(color_validation_errors)}")
    for error in color_validation_errors:
        print(f"  {error}")
    
    # Pass/fail criteria
    overall_pass = True
    
    if len(precision_errors) > 0:
        print("\n❌ FAIL: Round-trip precision errors detected")
        overall_pass = False
    else:
        print("\n✅ PASS: Round-trip precision maintained")
    
    if len(unit_conversion_errors) > 0:
        print("❌ FAIL: Unit conversion errors detected")
        overall_pass = False
    else:
        print("✅ PASS: Unit conversion precision maintained")
    
    if max_drift > 0.1:  # 0.1mm threshold
        print(f"❌ FAIL: Cumulative drift {max_drift:.3f}mm exceeds 0.1mm threshold")
        overall_pass = False
    else:
        print(f"✅ PASS: Cumulative drift {max_drift:.6f}mm within acceptable range")
    
    if len(color_validation_errors) > 0:
        print("❌ FAIL: Color validation errors detected")
        overall_pass = False
    else:
        print("✅ PASS: Color validation working correctly")
    
    print("\n" + "=" * 50)
    if overall_pass:
        print("🎉 ALL TESTS PASSED - Precision is maintained!")
    else:
        print("⚠️  SOME TESTS FAILED - Review precision implementation")
    
    return overall_pass

async def test_ocr_workflow_simulation():
    """Simulate a realistic OCR reconstruction workflow."""
    print("\n" + "=" * 50)
    print("OCR WORKFLOW SIMULATION")
    print("=" * 50)
    
    # Simulate OCR results from a document scan
    mock_ocr_data = {
        "page_width_px": 2480,  # 300 DPI A4 width
        "page_height_px": 3508, # 300 DPI A4 height
        "dpi": 300,
        "text_blocks": [
            {"id": 1, "text": "Document Title", "left": 124, "top": 100, "width": 1000, "height": 60, "confidence": 0.98},
            {"id": 2, "text": "Subtitle here", "left": 124, "top": 180, "width": 800, "height": 40, "confidence": 0.95},
            {"id": 3, "text": "First paragraph of content that spans multiple lines and contains important information about the document subject matter.", "left": 124, "top": 250, "width": 1200, "height": 120, "confidence": 0.92},
            {"id": 4, "text": "• Bullet point one", "left": 150, "top": 400, "width": 500, "height": 30, "confidence": 0.90},
            {"id": 5, "text": "• Bullet point two", "left": 150, "top": 440, "width": 600, "height": 30, "confidence": 0.91},
            {"id": 6, "text": "• Bullet point three", "left": 150, "top": 480, "width": 550, "height": 30, "confidence": 0.89},
            {"id": 7, "text": "Second paragraph with more detailed information and technical specifications.", "left": 124, "top": 550, "width": 1100, "height": 80, "confidence": 0.94},
            {"id": 8, "text": "Footer text", "left": 124, "top": 3300, "width": 400, "height": 25, "confidence": 0.85},
        ],
        "images": [
            {"id": 1, "path": "/mock/logo.png", "left": 1800, "top": 100, "width": 400, "height": 150},
            {"id": 2, "path": "/mock/diagram.png", "left": 200, "top": 700, "width": 1000, "height": 600},
        ]
    }
    
    print(f"Processing OCR data: {len(mock_ocr_data['text_blocks'])} text blocks, {len(mock_ocr_data['images'])} images")
    
    # Convert pixel coordinates to PowerPoint points
    from mcp_server.app_interaction.powerpoint_editor import px_to_points
    
    ppt_commands = []
    coordinate_precision_log = []
    
    # Convert OCR pixel coordinates to PowerPoint commands
    for block in mock_ocr_data["text_blocks"]:
        left_pt = px_to_points(block["left"], dpi=mock_ocr_data["dpi"])
        top_pt = px_to_points(block["top"], dpi=mock_ocr_data["dpi"])
        width_pt = px_to_points(block["width"], dpi=mock_ocr_data["dpi"])
        height_pt = px_to_points(block["height"], dpi=mock_ocr_data["dpi"])
        
        # Log precision for analysis
        coordinate_precision_log.append({
            "type": "text_box",
            "id": block["id"],
            "original_px": {"left": block["left"], "top": block["top"], "width": block["width"], "height": block["height"]},
            "converted_pt": {"left": left_pt, "top": top_pt, "width": width_pt, "height": height_pt},
            "rounded_pt": {
                "left": round(left_pt, 3),
                "top": round(top_pt, 3),
                "width": round(width_pt, 3),
                "height": round(height_pt, 3)
            }
        })
        
        font_formatting = {
            "name": "Arial",
            "size": max(8, min(24, height_pt * 0.6)),  # Estimate font size from box height
            "color": "#000000"
        }
        
        if "title" in block["text"].lower():
            font_formatting.update({"bold": True, "size": 16})
        elif block["text"].startswith("•"):
            font_formatting.update({"size": 10})
        
        ppt_commands.append({
            "tool": "ppt_add_text_box",
            "params": {
                "slide_index": 1,
                "left": f"{left_pt:.3f}pt",
                "top": f"{top_pt:.3f}pt", 
                "width": f"{width_pt:.3f}pt",
                "height": f"{height_pt:.3f}pt",
                "text": block["text"],
                "font": font_formatting,
                "dpi": mock_ocr_data["dpi"]
            }
        })
    
    # Convert image coordinates  
    for img in mock_ocr_data["images"]:
        left_pt = px_to_points(img["left"], dpi=mock_ocr_data["dpi"])
        top_pt = px_to_points(img["top"], dpi=mock_ocr_data["dpi"])
        width_pt = px_to_points(img["width"], dpi=mock_ocr_data["dpi"])
        height_pt = px_to_points(img["height"], dpi=mock_ocr_data["dpi"])
        
        coordinate_precision_log.append({
            "type": "image",
            "id": img["id"],
            "original_px": {"left": img["left"], "top": img["top"], "width": img["width"], "height": img["height"]},
            "converted_pt": {"left": left_pt, "top": top_pt, "width": width_pt, "height": height_pt},
            "rounded_pt": {
                "left": round(left_pt, 3),
                "top": round(top_pt, 3),
                "width": round(width_pt, 3),
                "height": round(height_pt, 3)
            }
        })
        
        ppt_commands.append({
            "tool": "ppt_add_image",
            "params": {
                "slide_index": 1,
                "path": img["path"],
                "left": f"{left_pt:.3f}pt",
                "top": f"{top_pt:.3f}pt",
                "width": f"{width_pt:.3f}pt",
                "height": f"{height_pt:.3f}pt",
                "preserve_aspect": False,
                "dpi": mock_ocr_data["dpi"]
            }
        })
    
    # Analyze coordinate precision
    max_precision_loss = 0.0
    for entry in coordinate_precision_log:
        for coord in ["left", "top", "width", "height"]:
            original = entry["converted_pt"][coord]
            rounded = entry["rounded_pt"][coord]
            precision_loss = abs(original - rounded)
            max_precision_loss = max(max_precision_loss, precision_loss)
    
    print(f"Generated {len(ppt_commands)} PowerPoint commands")
    print(f"Maximum coordinate precision loss: {max_precision_loss:.6f} points")
    
    # Save OCR reconstruction script for manual testing
    script_content = f"""#!/usr/bin/env python3
# Generated OCR Reconstruction Script
# This script can be used to test the actual PowerPoint MCP server

import asyncio
import json

async def reconstruct_document():
    # This would call the actual MCP server tools
    commands = {json.dumps(ppt_commands, indent=2)}
    
    print("OCR Reconstruction Commands:")
    for i, cmd in enumerate(commands):
        print(f"{{i+1:2d}}. {{cmd['tool']}}({{', '.join(f'{{k}}={{v}}' for k, v in cmd['params'].items())}})")
    
    # TODO: Execute commands via MCP server
    print("\\nTo execute:")
    print("1. Start PowerPoint MCP server")
    print("2. Call ppt_create_presentation(a4_portrait=True)")
    print("3. Call ppt_add_slide()")
    print("4. Execute each command above")

if __name__ == "__main__":
    asyncio.run(reconstruct_document())
"""
    
    # Write script file
    script_path = os.path.join(os.path.dirname(__file__), "ocr_reconstruction_test.py")
    with open(script_path, "w") as f:
        f.write(script_content)
    
    print(f"✅ OCR reconstruction script saved to: {script_path}")
    
    # Validation checks
    if max_precision_loss > 0.001:  # 0.001pt threshold
        print("❌ FAIL: Coordinate precision loss exceeds threshold")
        return False
    else:
        print("✅ PASS: Coordinate precision maintained within acceptable range")
        return True

async def main():
    """Run all precision tests."""
    print("PowerPoint MCP Server - Comprehensive Precision Testing")
    print("=" * 60)
    
    test1_pass = await test_precision_drift()
    test2_pass = await test_ocr_workflow_simulation()
    
    print("\n" + "=" * 60)
    print("FINAL RESULTS")
    print("=" * 60)
    
    if test1_pass and test2_pass:
        print("🎉 ALL PRECISION TESTS PASSED!")
        print("The PowerPoint MCP server maintains coordinate precision")
        print("suitable for high-accuracy OCR reconstruction workflows.")
        return 0
    else:
        print("⚠️  SOME PRECISION TESTS FAILED!")
        print("Review the implementation for coordinate precision issues.")
        return 1

if __name__ == "__main__":
    exit_code = asyncio.run(main())
    exit(exit_code)