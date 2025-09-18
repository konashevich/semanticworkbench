#!/usr/bin/env python3
# Generated OCR Reconstruction Script
# This script can be used to test the actual PowerPoint MCP server

import asyncio
import json

async def reconstruct_document():
    # This would call the actual MCP server tools
    commands = [
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "29.760pt",
                "top": "24.000pt",
                "width": "240.000pt",
                "height": "14.400pt",
                "text": "Document Title",
                "font": {"name": "Arial", "size": 16, "color": "#000000", "bold": True},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "29.760pt",
                "top": "43.200pt",
                "width": "192.000pt",
                "height": "9.600pt",
                "text": "Subtitle here",
                "font": {"name": "Arial", "size": 16, "color": "#000000", "bold": True},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "29.760pt",
                "top": "60.000pt",
                "width": "288.000pt",
                "height": "28.800pt",
                "text": "First paragraph of content that spans multiple lines and contains important information about the document subject matter.",
                "font": {"name": "Arial", "size": 17.28, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "36.000pt",
                "top": "96.000pt",
                "width": "120.000pt",
                "height": "7.200pt",
                "text": "\u2022 Bullet point one",
                "font": {"name": "Arial", "size": 10, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "36.000pt",
                "top": "105.600pt",
                "width": "144.000pt",
                "height": "7.200pt",
                "text": "\u2022 Bullet point two",
                "font": {"name": "Arial", "size": 10, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "36.000pt",
                "top": "115.200pt",
                "width": "132.000pt",
                "height": "7.200pt",
                "text": "\u2022 Bullet point three",
                "font": {"name": "Arial", "size": 10, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "29.760pt",
                "top": "132.000pt",
                "width": "264.000pt",
                "height": "19.200pt",
                "text": "Second paragraph with more detailed information and technical specifications.",
                "font": {"name": "Arial", "size": 11.52, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.textbox.add",
            "params": {
                "slide_index": 1,
                "left": "29.760pt",
                "top": "792.000pt",
                "width": "96.000pt",
                "height": "6.000pt",
                "text": "Footer text",
                "font": {"name": "Arial", "size": 8, "color": "#000000"},
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.image.add",
            "params": {
                "slide_index": 1,
                "path": "/mock/logo.png",
                "left": "432.000pt",
                "top": "24.000pt",
                "width": "96.000pt",
                "height": "36.000pt",
                "preserve_aspect": False,
                "dpi": 300,
            },
        },
        {
            "tool": "ppt.shape.image.add",
            "params": {
                "slide_index": 1,
                "path": "/mock/diagram.png",
                "left": "48.000pt",
                "top": "168.000pt",
                "width": "240.000pt",
                "height": "144.000pt",
                "preserve_aspect": False,
                "dpi": 300,
            },
        },
    ]

    print("OCR Reconstruction Commands:")
    for i, cmd in enumerate(commands):
        params_str = ", ".join(f"{k}={v}" for k, v in cmd["params"].items())
        print(f"{i+1:2d}. {cmd['tool']}({params_str})")

    print("\nTo execute:")
    print("1. Start PowerPoint MCP server")
    print("2. Call ppt.presentation.create(a4_portrait=True)")
    print("3. (Optional) Add more slides with ppt.slide.add()")
    print("4. Execute each command above")
