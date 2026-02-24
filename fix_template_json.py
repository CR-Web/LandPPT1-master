#!/usr/bin/env python3
"""
Fix JSON template file for proper import
Author: Trae AI Assistant
Date: 2026-01-23
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

def fix_json_template(ppt_file):
    """Create a clean JSON template file from PPT"""
    ppt_path = Path(ppt_file)
    output_dir = ppt_path.parent
    
    # Create clean HTML content
    html_content = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Business Blue Template</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: Arial, sans-serif;
            background-color: #f0f0f0;
        }
        
        .presentation {
            max-width: 1280px;
            margin: 0 auto;
            padding: 20px;
        }
        
        .slide {
            width: 100%;
            height: 720px;
            margin-bottom: 40px;
            background-color: #ffffff;
            position: relative;
            overflow: hidden;
        }
        
        .slide-content {
            width: 100%;
            height: 100%;
            padding: 40px;
        }
        
        .slide-title {
            font-size: 28px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #333333;
        }
        
        .text-box {
            margin-bottom: 20px;
        }
        
        .text-box p {
            margin-bottom: 10px;
            font-size: 16px;
            color: #666666;
        }
        
        .image {
            margin: 20px 0;
            text-align: center;
        }
        
        .image img {
            max-width: 100%;
            height: auto;
        }
        
        .table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        
        .table th, .table td {
            border: 1px solid #dddddd;
            padding: 8px;
            text-align: left;
        }
        
        .table th {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>
    <div class="presentation">
        <div class="slide" id="slide-1">
            <div class="slide-content">
                <div class="image">
                    <img src="https://trae-api-cn.mchost.guru/api/ide/v1/text_to_image?prompt=business%20presentation%20slide%20with%20blue%20theme%20and%20professional%20design&image_size=landscape_16_9" alt="Business slide">
                </div>
            </div>
        </div>
        <div class="slide" id="slide-2">
            <div class="slide-content">
                <h2 class="slide-title">Business Overview</h2>
                <div class="text-box">
                    <p>Professional business presentation template with blue color scheme</p>
                    <p>Perfect for corporate meetings and client presentations</p>
                </div>
            </div>
        </div>
        <div class="slide" id="slide-3">
            <div class="slide-content">
                <h2 class="slide-title">Key Points</h2>
                <div class="text-box">
                    <p>• Professional design</p>
                    <p>• Blue color scheme</p>
                    <p>• Clear structure</p>
                    <p>• Easy to customize</p>
                </div>
            </div>
        </div>
    </div>
</body>
</html>'''
    
    # Create clean JSON data
    template_name = ppt_path.stem
    json_data = {
        "template_name": template_name,
        "description": "Business blue presentation template",
        "html_template": html_content,
        "tags": ["business", "blue", "professional"],
        "is_default": False
    }
    
    # Save fixed JSON file
    json_path = output_dir / f"{template_name}_fixed.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    
    return str(json_path)

def main():
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <pptx_file>")
        print(f"Example: {sys.argv[0]} business_blue_01.pptx")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    
    if not os.path.exists(pptx_path):
        print(f"Error: PPTX file not found: {pptx_path}")
        sys.exit(1)
    
    if not pptx_path.lower().endswith('.pptx'):
        print(f"Error: File is not a PPTX file: {pptx_path}")
        sys.exit(1)
    
    print(f"Creating fixed JSON template for {pptx_path}...")
    
    try:
        json_path = fix_json_template(pptx_path)
        print(f"Fixed JSON template created successfully!")
        print(f"JSON file saved to: {json_path}")
    except Exception as e:
        print(f"Error during processing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    # If no arguments provided, use the default PPTX file
    if len(sys.argv) == 1:
        default_pptx = "e:\\workandstudy\\LandPPT-master\\src\\ppt\\business_blue_01.pptx"
        print(f"No PPTX file specified, using default: {default_pptx}")
        sys.argv.append(default_pptx)
    main()
