#!/usr/bin/env python3
"""
Project Presentation Generator
-------------------------------
Generates HTML slide presentations from README.md files.

Usage:
    python generate_presentation.py [path/to/readme.md] [output.html]
    
If no arguments provided, looks for README.md in current directory
and outputs to presentation.html

Author: True North Data Strategies
"""

import os
import sys
import re
import json
from datetime import datetime
from pathlib import Path


def parse_readme(readme_path: str) -> dict:
    """
    Parse a README.md file and extract structured content for presentation.
    
    Returns a dictionary with:
        - title: Project name
        - tagline: Short description
        - sections: List of content sections
        - features: List of key features
        - tech_stack: Technology information
        - installation: Setup instructions
        - usage: Usage examples
        - status: Project status/badges
        - contact: Contact information
        - license: License info
    """
    
    with open(readme_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    parsed = {
        'title': '',
        'tagline': '',
        'sections': [],
        'features': [],
        'tech_stack': [],
        'installation': [],
        'usage': [],
        'status': '',
        'contact': {},
        'license': '',
        'raw_sections': {}
    }
    
    # Extract title - first H1
    title_match = re.search(r'^#\s+(.+)$', content, re.MULTILINE)
    if title_match:
        parsed['title'] = title_match.group(1).strip()
    
    # Extract tagline - first paragraph after title or blockquote
    tagline_match = re.search(r'^#\s+.+\n+(?:>\s*)?(.+?)(?:\n\n|\n#)', content, re.MULTILINE)
    if tagline_match:
        tagline = tagline_match.group(1).strip()
        # Clean up blockquote markers
        tagline = re.sub(r'^>\s*', '', tagline)
        parsed['tagline'] = tagline
    
    # Split content into sections by H2 headers
    section_pattern = r'##\s+(.+?)\n([\s\S]*?)(?=\n##\s+|\Z)'
    sections = re.findall(section_pattern, content)
    
    for section_title, section_content in sections:
        section_title = section_title.strip()
        section_content = section_content.strip()
        
        # Store raw section content
        parsed['raw_sections'][section_title.lower()] = section_content
        
        # Categorize sections
        title_lower = section_title.lower()
        
        if any(word in title_lower for word in ['feature', 'capability', 'what', 'highlight']):
            parsed['features'] = extract_list_items(section_content)
            
        elif any(word in title_lower for word in ['tech', 'stack', 'built', 'tool', 'technology']):
            parsed['tech_stack'] = extract_list_items(section_content)
            
        elif any(word in title_lower for word in ['install', 'setup', 'getting started', 'quick start']):
            parsed['installation'] = extract_list_items(section_content)
            
        elif any(word in title_lower for word in ['usage', 'example', 'how to', 'demo']):
            parsed['usage'] = extract_code_blocks(section_content)
            
        elif any(word in title_lower for word in ['status', 'roadmap', 'progress']):
            parsed['status'] = section_content
            
        elif any(word in title_lower for word in ['contact', 'author', 'contributor', 'team']):
            parsed['contact'] = extract_contact_info(section_content)
            
        elif any(word in title_lower for word in ['license']):
            parsed['license'] = section_content
        
        # Add to general sections list
        parsed['sections'].append({
            'title': section_title,
            'content': section_content,
            'items': extract_list_items(section_content)
        })
    
    return parsed


def extract_list_items(content: str) -> list:
    """Extract bullet points and numbered items from markdown content."""
    items = []
    
    # Match bullet points (-, *, +) and numbered lists
    pattern = r'^[\s]*[-*+]\s+(.+)$|^[\s]*\d+\.\s+(.+)$'
    matches = re.findall(pattern, content, re.MULTILINE)
    
    for match in matches:
        item = match[0] or match[1]
        if item:
            # Clean up markdown formatting
            item = re.sub(r'\*\*(.+?)\*\*', r'\1', item)  # Bold
            item = re.sub(r'\*(.+?)\*', r'\1', item)      # Italic
            item = re.sub(r'`(.+?)`', r'\1', item)        # Code
            items.append(item.strip())
    
    return items


def extract_code_blocks(content: str) -> list:
    """Extract code blocks from markdown content."""
    blocks = []
    pattern = r'```(\w+)?\n([\s\S]*?)```'
    matches = re.findall(pattern, content)
    
    for lang, code in matches:
        blocks.append({
            'language': lang or 'text',
            'code': code.strip()
        })
    
    return blocks


def extract_contact_info(content: str) -> dict:
    """Extract contact information from content."""
    contact = {}
    
    # Email
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', content)
    if email_match:
        contact['email'] = email_match.group(0)
    
    # GitHub
    github_match = re.search(r'github\.com/([\w-]+)', content, re.IGNORECASE)
    if github_match:
        contact['github'] = github_match.group(1)
    
    # Website
    website_match = re.search(r'https?://(?!github\.com)[\w\.-]+\.\w+[/\w\.-]*', content)
    if website_match:
        contact['website'] = website_match.group(0)
    
    # LinkedIn
    linkedin_match = re.search(r'linkedin\.com/in/([\w-]+)', content, re.IGNORECASE)
    if linkedin_match:
        contact['linkedin'] = linkedin_match.group(1)
    
    return contact


def generate_slides(parsed: dict, config: dict = None) -> list:
    """
    Generate slide content from parsed README data.
    
    Returns a list of slide dictionaries with:
        - type: Slide template type
        - title: Slide title
        - content: Slide-specific content
    """
    
    config = config or {}
    slides = []
    
    # Slide 1: Title
    slides.append({
        'type': 'title',
        'title': parsed['title'],
        'tagline': parsed['tagline'] or 'Project Overview',
        'date': config.get('date', datetime.now().strftime('%B %Y'))
    })
    
    # Slide 2: Overview/Problem (if available)
    overview_sections = ['overview', 'about', 'introduction', 'problem', 'motivation']
    for key in overview_sections:
        if key in parsed['raw_sections']:
            content = parsed['raw_sections'][key]
            items = extract_list_items(content)
            slides.append({
                'type': 'content',
                'title': key.title(),
                'icon': get_icon_for_section(key),
                'items': items[:6] if items else [content[:500]],
                'highlight': items[0] if items else None
            })
            break
    
    # Slide 3: Features
    if parsed['features']:
        slides.append({
            'type': 'features',
            'title': 'Key Features',
            'icon': '🎯',
            'items': parsed['features'][:8]
        })
    
    # Slide 4: Tech Stack
    if parsed['tech_stack']:
        slides.append({
            'type': 'grid',
            'title': 'Tech Stack',
            'icon': '🛠️',
            'items': parsed['tech_stack'][:9]
        })
    
    # Slide 5-N: Additional sections
    skip_sections = ['feature', 'tech', 'stack', 'install', 'setup', 'license', 
                     'overview', 'about', 'introduction', 'contact', 'author']
    
    for section in parsed['sections']:
        title_lower = section['title'].lower()
        if not any(skip in title_lower for skip in skip_sections):
            if section['items']:
                slides.append({
                    'type': 'content',
                    'title': section['title'],
                    'icon': get_icon_for_section(title_lower),
                    'items': section['items'][:6]
                })
    
    # Installation slide
    if parsed['installation']:
        slides.append({
            'type': 'steps',
            'title': 'Getting Started',
            'icon': '🚀',
            'items': parsed['installation'][:6]
        })
    
    # Contact/Closing slide
    slides.append({
        'type': 'closing',
        'title': parsed['title'],
        'tagline': parsed['tagline'],
        'contact': parsed['contact'],
        'license': parsed['license']
    })
    
    return slides


def get_icon_for_section(section_name: str) -> str:
    """Return an appropriate emoji icon for a section."""
    icon_map = {
        'overview': '📋',
        'about': '📋',
        'problem': '🎯',
        'solution': '💡',
        'feature': '⭐',
        'architecture': '🏗️',
        'api': '🔌',
        'database': '🗄️',
        'security': '🔒',
        'performance': '⚡',
        'testing': '🧪',
        'deployment': '🚀',
        'roadmap': '🗺️',
        'contributing': '🤝',
        'changelog': '📝',
        'faq': '❓',
        'usage': '📖',
        'config': '⚙️',
        'integration': '🔗',
        'status': '📊'
    }
    
    for key, icon in icon_map.items():
        if key in section_name:
            return icon
    
    return '📌'


def render_html(slides: list, config: dict = None) -> str:
    """
    Render slides to HTML presentation.
    """
    
    config = config or {}
    theme = config.get('theme', {})
    
    # Theme defaults
    primary_gradient = theme.get('primary_gradient', 'linear-gradient(135deg, #1e3c72 0%, #2a5298 100%)')
    accent_color = theme.get('accent_color', '#ffd700')
    secondary_color = theme.get('secondary_color', '#87ceeb')
    
    html_slides = []
    
    for i, slide in enumerate(slides):
        slide_html = render_slide(slide, i == 0)
        html_slides.append(slide_html)
    
    total_slides = len(slides)
    title = slides[0]['title'] if slides else 'Presentation'
    
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title} - Presentation</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        :root {{
            --primary-gradient: {primary_gradient};
            --accent-color: {accent_color};
            --secondary-color: {secondary_color};
        }}

        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: var(--primary-gradient);
            color: white;
            overflow: hidden;
        }}

        .presentation-container {{
            width: 100vw;
            height: 100vh;
            position: relative;
        }}

        .slide {{
            width: 100%;
            height: 100%;
            padding: 60px;
            display: none;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            position: absolute;
            top: 0;
            left: 0;
        }}

        .slide.active {{
            display: flex;
        }}

        .slide h1 {{
            font-size: 3.5rem;
            margin-bottom: 2rem;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
        }}

        .slide h2 {{
            font-size: 2.5rem;
            margin-bottom: 2rem;
            color: var(--accent-color);
        }}

        .slide h3 {{
            font-size: 2rem;
            margin-bottom: 1.5rem;
            color: var(--secondary-color);
        }}

        .slide p {{
            font-size: 1.3rem;
            line-height: 1.6;
            margin-bottom: 1rem;
            max-width: 900px;
        }}

        .slide ul {{
            font-size: 1.2rem;
            text-align: left;
            max-width: 800px;
            margin: 1rem 0;
            list-style: none;
        }}

        .slide li {{
            margin-bottom: 0.8rem;
            padding-left: 1.5rem;
            position: relative;
        }}

        .slide li::before {{
            content: "▸";
            position: absolute;
            left: 0;
            color: var(--accent-color);
        }}

        .stat-grid {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 3rem;
            margin: 2rem 0;
            max-width: 800px;
        }}

        .stat-box {{
            background: rgba(255, 255, 255, 0.1);
            padding: 2rem;
            border-radius: 15px;
            backdrop-filter: blur(10px);
        }}

        .stat-number {{
            font-size: 3rem;
            font-weight: bold;
            color: var(--accent-color);
            margin-bottom: 0.5rem;
        }}

        .feature-grid {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 2rem;
            margin: 2rem 0;
            max-width: 1000px;
        }}

        .feature-card {{
            background: rgba(255, 255, 255, 0.1);
            padding: 1.5rem;
            border-radius: 10px;
            backdrop-filter: blur(10px);
        }}

        .feature-icon {{
            font-size: 2.5rem;
            margin-bottom: 1rem;
        }}

        .steps-container {{
            display: flex;
            flex-direction: column;
            gap: 1rem;
            max-width: 700px;
            width: 100%;
        }}

        .step-item {{
            display: flex;
            align-items: center;
            gap: 1rem;
            background: rgba(255, 255, 255, 0.1);
            padding: 1rem 1.5rem;
            border-radius: 10px;
            text-align: left;
        }}

        .step-number {{
            background: var(--accent-color);
            color: #1e3c72;
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
            flex-shrink: 0;
        }}

        .contact-info {{
            background: rgba(255, 255, 255, 0.1);
            padding: 2rem;
            border-radius: 15px;
            margin: 2rem 0;
            max-width: 600px;
            backdrop-filter: blur(10px);
        }}

        .contact-info a {{
            color: var(--secondary-color);
            text-decoration: none;
        }}

        .contact-info a:hover {{
            text-decoration: underline;
        }}

        .motto {{
            font-size: 2rem;
            color: var(--accent-color);
            font-weight: bold;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.5);
            margin-top: 2rem;
        }}

        .navigation {{
            position: fixed;
            bottom: 30px;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1rem;
            z-index: 1000;
        }}

        .nav-btn {{
            background: rgba(255, 255, 255, 0.2);
            border: 2px solid white;
            color: white;
            padding: 0.8rem 1.5rem;
            border-radius: 25px;
            cursor: pointer;
            font-size: 1rem;
            transition: all 0.3s ease;
            backdrop-filter: blur(10px);
        }}

        .nav-btn:hover {{
            background: rgba(255, 255, 255, 0.3);
            transform: translateY(-2px);
        }}

        .slide-counter {{
            position: fixed;
            top: 30px;
            right: 30px;
            background: rgba(0, 0, 0, 0.3);
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 1rem;
            backdrop-filter: blur(10px);
        }}

        .code-block {{
            background: rgba(0, 0, 0, 0.3);
            padding: 1.5rem;
            border-radius: 10px;
            text-align: left;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 1rem;
            max-width: 800px;
            overflow-x: auto;
            margin: 1rem 0;
        }}

        @media (max-width: 768px) {{
            .slide {{
                padding: 30px;
            }}
            
            .slide h1 {{
                font-size: 2.5rem;
            }}
            
            .slide h2 {{
                font-size: 2rem;
            }}
            
            .feature-grid,
            .stat-grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <div class="presentation-container">
        <div class="slide-counter">
            <span id="current-slide">1</span> / <span id="total-slides">{total_slides}</span>
        </div>

        {"".join(html_slides)}

        <div class="navigation">
            <button class="nav-btn" onclick="previousSlide()">← Previous</button>
            <button class="nav-btn" onclick="nextSlide()">Next →</button>
        </div>
    </div>

    <script>
        let currentSlide = 0;
        const slides = document.querySelectorAll('.slide');
        const totalSlides = slides.length;

        document.getElementById('total-slides').textContent = totalSlides;

        function showSlide(n) {{
            slides[currentSlide].classList.remove('active');
            currentSlide = (n + totalSlides) % totalSlides;
            slides[currentSlide].classList.add('active');
            document.getElementById('current-slide').textContent = currentSlide + 1;
        }}

        function nextSlide() {{
            showSlide(currentSlide + 1);
        }}

        function previousSlide() {{
            showSlide(currentSlide - 1);
        }}

        // Keyboard navigation
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'ArrowRight' || e.key === ' ') {{
                nextSlide();
            }} else if (e.key === 'ArrowLeft') {{
                previousSlide();
            }}
        }});

        // Touch/swipe navigation for mobile
        let startX = 0;
        let endX = 0;

        document.addEventListener('touchstart', function(e) {{
            startX = e.changedTouches[0].screenX;
        }});

        document.addEventListener('touchend', function(e) {{
            endX = e.changedTouches[0].screenX;
            handleSwipe();
        }});

        function handleSwipe() {{
            const threshold = 50;
            if (startX - endX > threshold) {{
                nextSlide();
            }} else if (endX - startX > threshold) {{
                previousSlide();
            }}
        }}
    </script>
</body>
</html>'''
    
    return html


def render_slide(slide: dict, is_first: bool = False) -> str:
    """Render a single slide to HTML."""
    
    active_class = ' active' if is_first else ''
    slide_type = slide.get('type', 'content')
    
    if slide_type == 'title':
        return f'''
        <div class="slide{active_class}">
            <h1>{slide['title']}</h1>
            <p style="font-size: 1.8rem; margin-bottom: 2rem;">{slide['tagline']}</p>
            <p style="margin-top: 3rem; font-size: 1.2rem;">{slide.get('date', '')}</p>
        </div>'''
    
    elif slide_type == 'features':
        items = slide.get('items', [])
        # Create grid cards for features
        cards_html = ''
        for item in items[:9]:
            # Try to extract title and description from "Title - Description" or "Title: Description" format
            # The item has already been cleaned of markdown formatting
            match = re.match(r'^([^-:]+?)\s*[-:]\s*(.+)$', item)
            if match:
                title = match.group(1).strip()
                desc = match.group(2).strip()
            else:
                # If no separator, use the whole item as title
                title = item
                desc = ''
            
            # Truncate title if too long
            if len(title) > 50:
                title = title[:47] + '...'
            
            cards_html += f'''
                <div class="feature-card">
                    <h4>{title}</h4>
                    <p style="font-size: 1rem;">{desc}</p>
                </div>'''
        
        return f'''
        <div class="slide{active_class}">
            <h2>{slide.get('icon', '⭐')} {slide['title']}</h2>
            <div class="feature-grid">
                {cards_html}
            </div>
        </div>'''
    
    elif slide_type == 'grid':
        items = slide.get('items', [])
        cards_html = ''
        for item in items[:9]:
            cards_html += f'''
                <div class="feature-card">
                    <p>{item}</p>
                </div>'''
        
        return f'''
        <div class="slide{active_class}">
            <h2>{slide.get('icon', '📌')} {slide['title']}</h2>
            <div class="feature-grid">
                {cards_html}
            </div>
        </div>'''
    
    elif slide_type == 'steps':
        items = slide.get('items', [])
        steps_html = ''
        for i, item in enumerate(items[:6], 1):
            steps_html += f'''
                <div class="step-item">
                    <div class="step-number">{i}</div>
                    <span>{item}</span>
                </div>'''
        
        return f'''
        <div class="slide{active_class}">
            <h2>{slide.get('icon', '🚀')} {slide['title']}</h2>
            <div class="steps-container">
                {steps_html}
            </div>
        </div>'''
    
    elif slide_type == 'closing':
        contact = slide.get('contact', {})
        contact_html = ''
        
        if contact:
            contact_items = []
            if contact.get('email'):
                contact_items.append(f'<p><strong>Email:</strong> <a href="mailto:{contact["email"]}">{contact["email"]}</a></p>')
            if contact.get('github'):
                contact_items.append(f'<p><strong>GitHub:</strong> <a href="https://github.com/{contact["github"]}" target="_blank">github.com/{contact["github"]}</a></p>')
            if contact.get('website'):
                contact_items.append(f'<p><strong>Website:</strong> <a href="{contact["website"]}" target="_blank">{contact["website"]}</a></p>')
            if contact.get('linkedin'):
                contact_items.append(f'<p><strong>LinkedIn:</strong> <a href="https://linkedin.com/in/{contact["linkedin"]}" target="_blank">linkedin.com/in/{contact["linkedin"]}</a></p>')
            
            if contact_items:
                contact_html = f'''
                <div class="contact-info">
                    <h3>Contact</h3>
                    {"".join(contact_items)}
                </div>'''
        
        return f'''
        <div class="slide{active_class}">
            <h1>{slide['title']}</h1>
            <p style="font-size: 1.5rem;">{slide.get('tagline', '')}</p>
            {contact_html}
            <p style="margin-top: 2rem; font-size: 1.2rem;">Thank you!</p>
        </div>'''
    
    else:  # Default content slide
        items = slide.get('items', [])
        items_html = ''
        if items:
            items_html = '<ul>' + ''.join(f'<li>{item}</li>' for item in items) + '</ul>'
        
        return f'''
        <div class="slide{active_class}">
            <h2>{slide.get('icon', '📌')} {slide['title']}</h2>
            {items_html}
        </div>'''


def main():
    """Main entry point for the presentation generator."""
    
    # Parse command line arguments
    readme_path = 'README.md'
    output_path = 'presentation.html'
    config_path = None
    
    args = sys.argv[1:]
    
    if len(args) >= 1:
        readme_path = args[0]
    if len(args) >= 2:
        output_path = args[1]
    if len(args) >= 3:
        config_path = args[2]
    
    # Check if README exists
    if not os.path.exists(readme_path):
        print(f"Error: README file not found: {readme_path}")
        print("\nUsage: python generate_presentation.py [README.md] [output.html] [config.json]")
        sys.exit(1)
    
    # Load config if provided
    config = {}
    if config_path and os.path.exists(config_path):
        with open(config_path, 'r') as f:
            config = json.load(f)
    
    print(f"Parsing README: {readme_path}")
    parsed = parse_readme(readme_path)
    
    print(f"Generating slides for: {parsed['title']}")
    slides = generate_slides(parsed, config)
    
    print(f"Rendering {len(slides)} slides to HTML...")
    html = render_html(slides, config)
    
    # Write output
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"Presentation saved to: {output_path}")
    print(f"Open in a browser to view your presentation!")


if __name__ == '__main__':
    main()
