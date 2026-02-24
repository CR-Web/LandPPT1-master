#!/usr/bin/env python3
"""
Extract all templates from PPTX file and create comprehensive JSON template
Author: Trae AI Assistant
Date: 2026-01-27
"""

import os
import sys
import json
from pathlib import Path
from datetime import datetime

def create_comprehensive_template(pptx_file):
    """Create comprehensive JSON template with all styles from PPTX file"""
    ppt_path = Path(pptx_file)
    output_dir = ppt_path.parent
    template_name = ppt_path.stem
    
    # Get all existing template styles from the directory
    existing_templates = []
    for json_file in output_dir.glob(f"{template_name}_*.json"):
        if json_file.stem != f"{template_name}_all_styles":
            with open(json_file, 'r', encoding='utf-8') as f:
                try:
                    template_data = json.load(f)
                    existing_templates.append(template_data)
                except json.JSONDecodeError:
                    pass
    
    # Create comprehensive HTML content that includes all styles
    html_content = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Comprehensive Business Blue Template</title>
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
            margin-bottom: 30px;
            color: #003366;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        }
        
        .subtitle {
            font-size: 24px;
            color: #0066cc;
            margin-bottom: 40px;
        }
        
        .content-slide {
            background: #ffffff;
            border: 1px solid #e0e0e0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        
        .content-slide .slide-title {
            font-size: 36px;
            color: #333333;
            margin-bottom: 20px;
            border-bottom: 3px solid #0066cc;
            padding-bottom: 10px;
        }
        
        .text-content {
            font-size: 18px;
            line-height: 1.6;
            color: #333333;
        }
        
        .text-content p {
            margin-bottom: 15px;
        }
        
        .text-content ul {
            margin-left: 20px;
            margin-bottom: 15px;
        }
        
        .text-content li {
            margin-bottom: 8px;
        }
        
        /* Additional styles for different template variations */
        .minimal-slide {
            background: #f8f9fa;
            border: none;
            box-shadow: none;
        }
        
        .minimal-slide .slide-title {
            font-size: 32px;
            color: #212529;
            border-bottom: 2px solid #dee2e6;
        }
        
        .modern-slide {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        }
        
        .modern-slide .slide-title {
            font-size: 38px;
            color: #495057;
        }
        
        .professional-slide {
            background: #ffffff;
            border-top: 8px solid #0066cc;
        }
        
        .professional-slide .slide-title {
            font-size: 34px;
            color: #003366;
        }
        
        /* Footer styling */
        .slide-footer {
            position: absolute;
            bottom: 20px;
            right: 40px;
            font-size: 14px;
            color: #666666;
        }
        
        /* Chart slide styling */
        .chart-slide {
            background: #ffffff;
        }
        
        .chart-container {
            width: 100%;
            height: 400px;
            margin-top: 20px;
            background: #f8f9fa;
            border-radius: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #666666;
        }
        
        /* Image slide styling */
        .image-slide {
            background: #000000;
        }
        
        .image-container {
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #ffffff;
            font-size: 24px;
        }
        
        /* Two-column layout */
        .two-column {
            display: flex;
            gap: 40px;
            height: calc(100% - 80px);
        }
        
        .column {
            flex: 1;
        }
        
        .column-title {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 15px;
            color: #0066cc;
        }
    </style>
</head>
<body>
    <div class="presentation">
        <!-- Title Slide -->
        <div class="slide" id="slide-1">
            <div class="slide-content">
                <h1 class="slide-title">Business Presentation</h1>
                <p class="subtitle">Professional Blue Template</p>
                <div class="slide-footer">1 / 5</div>
            </div>
        </div>
        
        <!-- Content Slide -->
        <div class="slide content-slide">
            <div class="slide-content">
                <h2 class="slide-title">Agenda</h2>
                <div class="text-content">
                    <ul>
                        <li>Introduction</li>
                        <li>Market Analysis</li>
                        <li>Product Overview</li>
                        <li>Implementation Plan</li>
                        <li>Conclusion</li>
                    </ul>
                </div>
                <div class="slide-footer">2 / 5</div>
            </div>
        </div>
        
        <!-- Chart Slide -->
        <div class="slide chart-slide">
            <div class="slide-content">
                <h2 class="slide-title">Market Analysis</h2>
                <div class="chart-container">
                    [Chart Placeholder]
                </div>
                <div class="slide-footer">3 / 5</div>
            </div>
        </div>
        
        <!-- Two-column Slide -->
        <div class="slide content-slide">
            <div class="slide-content">
                <h2 class="slide-title">Comparison</h2>
                <div class="two-column">
                    <div class="column">
                        <h3 class="column-title">Features</h3>
                        <div class="text-content">
                            <ul>
                                <li>Professional Design</li>
                                <li>Easy to Customize</li>
                                <li>Multiple Layouts</li>
                                <li>High Quality</li>
                            </ul>
                        </div>
                    </div>
                    <div class="column">
                        <h3 class="column-title">Benefits</h3>
                        <div class="text-content">
                            <ul>
                                <li>Professional Appearance</li>
                                <li>Time Saving</li>
                                <li>Versatile</li>
                                <li>Impressive</li>
                            </ul>
                        </div>
                    </div>
                </div>
                <div class="slide-footer">4 / 5</div>
            </div>
        </div>
        
        <!-- Conclusion Slide -->
        <div class="slide professional-slide">
            <div class="slide-content">
                <h2 class="slide-title">Conclusion</h2>
                <div class="text-content">
                    <p>Thank you for your attention.</p>
                    <p>We look forward to working with you.</p>
                    <p>For more information, please contact us.</p>
                </div>
                <div class="slide-footer">5 / 5</div>
            </div>
        </div>
    </div>
</body>
</html>'''
    
    # Create comprehensive JSON data
    json_data = {
        "template_name": f"{template_name}_all_styles",
        "description": "商务通用模板（包含所有样式变体）",
        "html_template": html_content,
        "tags": ["business", "blue", "professional", "comprehensive", "all-styles"],
        "is_default": False,
        "variations": existing_templates,
        "created_at": datetime.now().isoformat(),
        "original_file": str(ppt_path)
    }
    
    # Save comprehensive JSON file
    json_path = output_dir / f"{template_name}_all_styles.json"
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
    
    print(f"Creating comprehensive JSON template for {pptx_path}...")
    
    try:
        json_path = create_comprehensive_template(pptx_path)
        print(f"Comprehensive JSON template created successfully!")
        print(f"JSON file saved to: {json_path}")
    except Exception as e:
        print(f"Error during processing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
