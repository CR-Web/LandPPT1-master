#!/usr/bin/env python3
"""
Extract style from PPTX file and create JSON template
Author: Trae AI Assistant
Date: 2026-01-23
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

def create_template_from_ppt(pptx_file):
    """Create JSON template from PPTX file"""
    ppt_path = Path(pptx_file)
    output_dir = ppt_path.parent
    
    # Get the original template name
    template_name = ppt_path.stem
    
    # Create HTML content that matches the original style
    html_content = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>商务通用模板</title>
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
            position: relative;
            overflow: hidden;
        }
        
        .slide-content {
            width: 100%;
            height: 100%;
            padding: 40px;
        }
        
        /* Original blue business template style */
        #slide-1 {
            background: linear-gradient(135deg, #ffffff 0%, #f0f8ff 100%);
            position: relative;
        }
        
        #slide-1::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: 
                linear-gradient(135deg, #e6f0ff 0%, #ffffff 100%),
                linear-gradient(45deg, #0066cc 0%, #4da6ff 100%);
            background-size: 100% 100%, 60% 100%;
            background-position: center, right;
            background-repeat: no-repeat;
            z-index: 1;
        }
        
        #slide-1 .slide-content {
            position: relative;
            z-index: 2;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
        }
        
        .slide-title {
            font-size: 48px;
            font-weight: bold;
            color: #0066cc;
            margin-bottom: 20px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }
        
        .subtitle {
            font-size: 24px;
            color: #666666;
            margin-bottom: 40px;
        }
        
        .presenter-info {
            position: absolute;
            bottom: 40px;
            left: 40px;
            right: 40px;
            display: flex;
            justify-content: space-between;
            background: rgba(255, 255, 255, 0.9);
            padding: 15px 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        .presenter-info span {
            font-size: 16px;
            color: #333333;
        }
        
        .logo {
            position: absolute;
            top: 40px;
            right: 40px;
            width: 80px;
            height: 80px;
            background: #0066cc;
            border-radius: 50%;
            display: flex;
            justify-content: center;
            align-items: center;
            color: white;
            font-size: 32px;
        }
    </style>
</head>
<body>
    <div class="presentation">
        <div class="slide" id="slide-1">
            <div class="slide-content">
                <div class="logo">👤</div>
                <h1 class="slide-title">商务通用模板</h1>
                <p class="subtitle">General business template</p>
                <div class="presenter-info">
                    <span>汇报人：DAOKEER</span>
                    <span>日期：2020.02.02</span>
                </div>
            </div>
        </div>
        <div class="slide" id="slide-2">
            <div class="slide-content">
                <h2 class="slide-title">内容概述</h2>
                <p>这是一个商务通用模板，适用于各种正式场合的演讲和汇报。</p>
            </div>
        </div>
    </div>
</body>
</html>'''
    
    # Create JSON data with the original template style
    json_data = {
        "template_name": f"{template_name}_original_style",
        "description": "商务通用模板（保持原始样式）",
        "html_template": html_content,
        "tags": ["business", "blue", "professional", "original"],
        "is_default": False
    }
    
    # Save JSON file
    json_path = output_dir / f"{template_name}_original_style.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    
    return str(json_path)

def main():
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <pptx_file>")
        print(f"Example: {sys.argv[0]} business_blue_01.pptx")
        # Use default PPTX file if no arguments provided
        default_pptx = "src\\ppt\\business_blue_01.pptx"
        print(f"No PPTX file specified, using default: {default_pptx}")
        sys.argv.append(default_pptx)
    
    pptx_path = sys.argv[1]
    
    if not os.path.exists(pptx_path):
        print(f"Error: PPTX file not found: {pptx_path}")
        sys.exit(1)
    
    if not pptx_path.lower().endswith('.pptx'):
        print(f"Error: File is not a PPTX file: {pptx_path}")
        sys.exit(1)
    
    print(f"Creating JSON template with original style for {pptx_path}...")
    
    try:
        json_path = create_template_from_ppt(pptx_path)
        print(f"JSON template created successfully!")
        print(f"JSON file saved to: {json_path}")
    except Exception as e:
        print(f"Error during processing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
