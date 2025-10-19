from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import openai
import os
from dotenv import load_dotenv
import json
import requests
from pydub import AudioSegment
from PIL import Image, ImageDraw, ImageFont
import cv2
import numpy as np
from moviepy.editor import ImageClip, AudioFileClip, concatenate_videoclips
import pickle
from datetime import datetime

load_dotenv()

client = openai.OpenAI(api_key=os.getenv("OPENAI_KEY"))

LEARNING_FILE = "mermaid_patterns.pkl"

def load_successful_patterns():
    if os.path.exists(LEARNING_FILE):
        try:
            with open(LEARNING_FILE, 'rb') as f:
                return pickle.load(f)
        except:
            pass
    return {"successful_patterns": [], "failed_patterns": []}

def save_successful_pattern(pattern, success=True):
    data = load_successful_patterns()
    
    if success:
        data["successful_patterns"].append({
            "pattern": pattern,
            "timestamp": datetime.now().isoformat()
        })
        data["successful_patterns"] = data["successful_patterns"][-50:]
    else:
        data["failed_patterns"].append({
            "pattern": pattern,
            "timestamp": datetime.now().isoformat()
        })
        data["failed_patterns"] = data["failed_patterns"][-20:]
    
    with open(LEARNING_FILE, 'wb') as f:
        pickle.dump(data, f)

def get_pattern_examples():
    """Get examples from successful patterns"""
    data = load_successful_patterns()
    if data["successful_patterns"]:
        # Return last 5 successful patterns as examples
        examples = data["successful_patterns"][-5:]
        return "\n\nEXAMPLES OF SUCCESSFUL PATTERNS:\n" + "\n---\n".join(
            [p["pattern"] for p in examples]
        )
    return ""

def generate_slide_content(topic, num_slides=5):
    """Generate content for slides using GPT with structured JSON output"""
    prompt = f"""Create content for a {num_slides}-slide presentation about {topic}.
    
    Return a JSON object with this exact structure:
    {{
        "slides": [
            {{
                "slide_number": 1,
                "type": "title",
                "title": "Presentation Title",
                "subtitle": "Optional Subtitle"
            }},
            {{
                "slide_number": 2,
                "type": "content",
                "title": "Slide Title",
                "bullets": [
                    "Bullet point 1",
                    "Bullet point 2",
                    "Bullet point 3"
                ]
            }}
        ]
    }}
    
    Generate {num_slides} slides total. First slide must be type "title", rest are type "content"."""
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a presentation content creator. Create clear, concise bullet points. Always respond with valid JSON only."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"},
        temperature=0.7
    )
    
    return response.choices[0].message.content

def generate_mermaid_chart(slide_data, topic, max_retries=10):
    """Generate a complex Mermaid chart with multiple arrows using iterative learning"""
    
    bullets = '\n'.join(slide_data.get('bullets', []))
    
    pattern_examples = get_pattern_examples()
    
    prompt = f"""Create a Mermaid flowchart with MULTIPLE ARROWS and COMPLEX RELATIONSHIPS:

Title: {slide_data.get('title')}
Content: {bullets}

CRITICAL SYNTAX RULES (FOLLOW EXACTLY):
1. Start with EXACTLY: %%{{init: {{'theme':'forest'}}}}%%
2. Second line MUST be: graph TD
3. Node IDs: Use ONLY A, B, C, D, E, F, G, H (single letters)
4. Node text: A[Text Here] - NO quotes, NO special chars in text
5. Arrows: Use --> ONLY
6. MULTIPLE ARROWS: Each node can connect to multiple nodes
   Example: A --> B
            A --> C
            B --> D
            C --> D
7. Create 5-8 nodes with COMPLEX relationships
8. NO subgraphs, NO styles, NO classes, NO special syntax
9. Each line: one connection only

GOOD COMPLEX EXAMPLE:
%%{{init: {{'theme':'forest'}}}}%%
graph TD
    A[Start] --> B[Process 1]
    A --> C[Process 2]
    B --> D[Result 1]
    C --> D[Result 1]
    B --> E[Result 2]
    C --> F[Result 3]
    D --> G[Conclusion]
    E --> G[Conclusion]
    F --> G[Conclusion]

BAD EXAMPLES TO AVOID:
- Using quotes: A["Text"]
- Special chars: A[Text!@#]
- Multi-word IDs: Node1[Text]
- Multiple arrows per line: A --> B --> C
- Missing init line
- Wrong arrow: A -> B or A => B

{pattern_examples}

Create a COMPLEX chart with MULTIPLE ARROWS. Return ONLY the Mermaid code, nothing else."""
    
    for attempt in range(max_retries):
        try:
            print(f"  🔄 Attempt {attempt + 1}/{max_retries} to generate valid Mermaid chart...")
            
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a Mermaid diagram expert. Create ONLY valid, complex flowcharts with multiple arrows. Follow syntax rules EXACTLY."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2  # Low temperature for consistency
            )
            
            mermaid_code = response.choices[0].message.content.strip()
            
            # Clean up markdown code blocks
            if "```mermaid" in mermaid_code:
                mermaid_code = mermaid_code.split("```mermaid")[1].split("```")[0].strip()
            elif "```" in mermaid_code:
                mermaid_code = mermaid_code.replace("```", "").strip()
            
            # Strict validation
            lines = mermaid_code.split('\n')
            
            # Check init line
            if not lines[0].strip().startswith("%%{init:"):
                print(f"  ❌ Missing init line")
                save_successful_pattern(mermaid_code, success=False)
                continue
            
            # Check graph declaration
            if "graph TD" not in lines[1] and "graph LR" not in lines[1]:
                print(f"  ❌ Missing or incorrect graph declaration")
                save_successful_pattern(mermaid_code, success=False)
                continue
            
            # Validate arrows
            arrow_count = mermaid_code.count('-->')
            if arrow_count < 4:
                print(f"  ❌ Not enough complexity ({arrow_count} arrows, need 4+)")
                save_successful_pattern(mermaid_code, success=False)
                continue
            
            # Check for bad patterns
            if '["' in mermaid_code or "['" in mermaid_code:
                print(f"  ❌ Contains quotes in brackets")
                save_successful_pattern(mermaid_code, success=False)
                continue
            
            # Test render
            if validate_mermaid_syntax(mermaid_code):
                print(f"  ✅ Valid complex Mermaid chart generated with {arrow_count} arrows")
                save_successful_pattern(mermaid_code, success=True)
                return mermaid_code
            else:
                print(f"  ❌ Syntax validation failed on render test")
                save_successful_pattern(mermaid_code, success=False)
                
                # On later attempts, provide more specific feedback
                if attempt >= 3:
                    prompt += f"\n\nPREVIOUS ATTEMPT FAILED. Common issues:\n- Check all brackets are properly closed\n- Ensure no special characters\n- Verify arrow syntax (use --> only)\n- Make sure each connection is on its own line"
                
        except Exception as e:
            print(f"  ❌ Error in attempt {attempt + 1}: {e}")
            if attempt < max_retries - 1:
                continue
    
    # If all attempts fail, raise an error instead of using fallback
    raise Exception(f"Failed to generate valid Mermaid chart after {max_retries} attempts. Check the learning file and try again.")

def validate_mermaid_syntax(mermaid_code):
    """Validate Mermaid syntax by attempting to render it"""
    try:
        import base64
        encoded = base64.b64encode(mermaid_code.encode('utf-8')).decode('utf-8')
        url = f"https://mermaid.ink/img/{encoded}"
        response = requests.get(url, timeout=15)
        
        # Check if it's actually an image
        if response.status_code == 200:
            content_type = response.headers.get('content-type', '')
            return 'image' in content_type and len(response.content) > 1000
        return False
    except:
        return False

def render_mermaid_to_image(mermaid_code, output_path, max_retries=3):
    """Render Mermaid code to image using Mermaid.ink API"""
    import base64
    import time
    
    for attempt in range(max_retries):
        try:
            clean_code = mermaid_code.strip()
            encoded = base64.b64encode(clean_code.encode('utf-8')).decode('utf-8')
            url = f"https://mermaid.ink/img/{encoded}"
            
            print(f"  📊 Rendering Mermaid chart to image (attempt {attempt + 1}/{max_retries})...")
            response = requests.get(url, timeout=30)
            
            if response.status_code == 200:
                content_type = response.headers.get('content-type', '')
                if 'image' in content_type:
                    with open(output_path, 'wb') as f:
                        f.write(response.content)
                    print(f"  ✅ Mermaid chart rendered successfully")
                    return True
                else:
                    print(f"  ⚠️  Response is not an image")
                    time.sleep(2)
                    continue
            else:
                print(f"  ❌ HTTP {response.status_code}")
                time.sleep(2)
                
        except Exception as e:
            print(f"  ❌ Render error: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
    
    raise Exception("Failed to render Mermaid chart to image")

def generate_narration_script(slide_data):
    """Generate natural speech narration for a slide using GPT"""
    if slide_data.get('type') == 'title':
        prompt = f"""Create a brief spoken introduction (2-3 sentences) for this presentation title slide. Script should not be more than 30 seconds long.

        Title: {slide_data.get('title')}
        Subtitle: {slide_data.get('subtitle', '')}
        
        Write it as natural speech, as if presenting to an audience. Keep it concise and welcoming."""
    else:
        bullets = '\n'.join(f"- {b}" for b in slide_data.get('bullets', []))
        prompt = f"""Create natural spoken narration for this slide. Explain each point conversationally:
        
        Slide Title: {slide_data.get('title')}
        Points:
        {bullets}
        DO NOT Start with Hello Everyone or welcome because this is not the first slide.
        talk like a continuation of the presentation in the beginning.
        Script should not be more than 30 seconds long.
        Write as if you're presenting to an audience. Be clear and engaging. Don't just read the bullets - explain them naturally."""
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a professional presenter. Create natural, engaging spoken narration."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.7
    )
    
    return response.choices[0].message.content

def generate_speech_elevenlabs(text, output_path, voice_id=None, api_key=None):
    """Generate speech using ElevenLabs API and return audio duration in seconds"""
    if voice_id is None:
        voice_id = os.getenv("ELEVENLABS_VOICE_ID")
    if api_key is None:
        api_key = os.getenv("ELEVENLABS_API_KEY")
    
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
    
    headers = {
        "Accept": "audio/mpeg",
        "Content-Type": "application/json",
        "xi-api-key": api_key
    }
    
    data = {
        "text": text,
        "model_id": "eleven_monolingual_v1",
        "voice_settings": {
            "stability": 0.5,
            "similarity_boost": 0.75
        }
    }
    
    print(f"  🎤 Generating speech...")
    response = requests.post(url, json=data, headers=headers)
    
    if response.status_code == 200:
        with open(output_path, 'wb') as f:
            f.write(response.content)
        
        audio = AudioSegment.from_mp3(output_path)
        duration_seconds = len(audio) / 1000.0
        
        print(f"  ✅ Audio saved: {output_path} (Duration: {duration_seconds:.1f}s)")
        return duration_seconds
    else:
        print(f"  ❌ Error generating speech: {response.status_code}")
        print(f"  Response: {response.text}")
        return 5.0

def create_slide_image(slide_data, logo_path, output_path, mermaid_image_path=None, 
                      width=1920, height=1080):
    """Create a slide image using PIL with LARGER Mermaid chart"""
    img = Image.new('RGB', (width, height), color='white')
    draw = ImageDraw.Draw(img)
    
    # Load logo
    try:
        logo = Image.open(logo_path)
        logo_height = 80
        aspect = logo.width / logo.height
        logo_width = int(logo_height * aspect)
        logo = logo.resize((logo_width, logo_height), Image.Resampling.LANCZOS)
        img.paste(logo, (50, height - logo_height - 50), logo if logo.mode == 'RGBA' else None)
    except Exception as e:
        print(f"  ⚠️  Could not load logo: {e}")
    
    # Load fonts
    try:
        title_font = ImageFont.truetype("Arial", 80)
        subtitle_font = ImageFont.truetype("Arial", 40)
        content_title_font = ImageFont.truetype("Arial", 60)
        bullet_font = ImageFont.truetype("Arial", 35)
    except:
        try:
            title_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 80)
            subtitle_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 40)
            content_title_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 60)
            bullet_font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 35)
        except:
            title_font = ImageFont.load_default()
            subtitle_font = ImageFont.load_default()
            content_title_font = ImageFont.load_default()
            bullet_font = ImageFont.load_default()
    
    if slide_data.get('type') == 'title':
        # Title slide - NO CHART
        title = slide_data.get('title', 'Presentation Title')
        subtitle = slide_data.get('subtitle', '')
        
        title_bbox = draw.textbbox((0, 0), title, font=title_font)
        title_width = title_bbox[2] - title_bbox[0]
        title_x = (width - title_width) // 2
        draw.text((title_x, height // 2 - 100), title, fill=(31, 73, 125), font=title_font)
        
        if subtitle:
            subtitle_bbox = draw.textbbox((0, 0), subtitle, font=subtitle_font)
            subtitle_width = subtitle_bbox[2] - subtitle_bbox[0]
            subtitle_x = (width - subtitle_width) // 2
            draw.text((subtitle_x, height // 2 + 50), subtitle, fill=(89, 89, 89), font=subtitle_font)
    
    else:
        # Content slide with LARGER chart
        title = slide_data.get('title', 'Slide Title')
        bullets = slide_data.get('bullets', [])
        
        draw.text((150, 80), title, fill=(31, 73, 125), font=content_title_font)
        
        # Draw bullets (reduce max width to make room for larger chart)
        y_position = 250
        max_bullet_width = 800  # Reduced from default to make room for chart
        
        for bullet in bullets:
            words = bullet.split()
            lines = []
            current_line = []
            
            for word in words:
                current_line.append(word)
                test_line = ' '.join(current_line)
                bbox = draw.textbbox((0, 0), test_line, font=bullet_font)
                if bbox[2] - bbox[0] > max_bullet_width:
                    current_line.pop()
                    lines.append(' '.join(current_line))
                    current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))
            
            draw.ellipse([150, y_position + 10, 165, y_position + 25], fill=(31, 73, 125))
            
            for line in lines:
                draw.text((200, y_position), line, fill=(0, 0, 0), font=bullet_font)
                y_position += 50
            
            y_position += 20
    
    # Add LARGER Mermaid chart for content slides
    if mermaid_image_path and os.path.exists(mermaid_image_path):
        try:
            mermaid_img = Image.open(mermaid_image_path)
            
            # MUCH LARGER chart dimensions
            max_chart_width = 800  # Increased from 500
            max_chart_height = 600  # Increased from 350
            
            aspect = mermaid_img.width / mermaid_img.height
            
            if mermaid_img.width > max_chart_width or mermaid_img.height > max_chart_height:
                if aspect > 1:
                    new_width = max_chart_width
                    new_height = int(new_width / aspect)
                else:
                    new_height = max_chart_height
                    new_width = int(new_height * aspect)
                
                mermaid_img = mermaid_img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            
            # Position in right side with padding
            chart_x = width - mermaid_img.width - 50
            chart_y = (height - mermaid_img.height) // 2  # Centered vertically
            
            if mermaid_img.mode == 'RGBA':
                img.paste(mermaid_img, (chart_x, chart_y), mermaid_img)
            else:
                img.paste(mermaid_img, (chart_x, chart_y))
                
            print(f"  📊 LARGE Mermaid chart added to slide ({mermaid_img.width}x{mermaid_img.height})")
        except Exception as e:
            print(f"  ⚠️  Could not add Mermaid chart to slide: {e}")
    
    img.save(output_path)
    print(f"  🖼️  Slide image saved: {output_path}")

def parse_gpt_content(content):
    """Parse GPT JSON response into structured slide data"""
    data = json.loads(content)
    return data['slides']

def generate_video_presentation(topic, logo_path="logo.png", output_filename=None, 
                                audio_dir="audio", slides_dir="slides", charts_dir="charts"):
    """
    Generate complete video presentation with narration and dynamic Mermaid charts.
    - NO chart on title slide
    - Complex charts with multiple arrows on content slides
    - Iterative learning to fix syntax errors
    """
    os.makedirs(audio_dir, exist_ok=True)
    os.makedirs(slides_dir, exist_ok=True)
    os.makedirs(charts_dir, exist_ok=True)
    
    if output_filename is None:
        safe_topic = "".join(c if c.isalnum() or c in (' ', '-', '_') else '' for c in topic)
        safe_topic = safe_topic.replace(' ', '_')[:50]
        output_filename = f"{safe_topic}_presentation.mp4"
    
    print(f"🎨 Generating presentation about: {topic}")
    print("⏳ Creating content with GPT...")
    content = generate_slide_content(topic)
    
    print("📝 Parsing structured content...")
    slides_data = parse_gpt_content(content)
    print(f"✅ Generated {len(slides_data)} slides")
    
    video_clips = []
    
    for idx, slide_data in enumerate(slides_data, 1):
        print(f"\n📄 Processing Slide {idx}/{len(slides_data)}")
        
        mermaid_image_path = None
        
        # Only generate Mermaid chart for CONTENT slides (not title)
        if slide_data.get('type') != 'title':
            print(f"  📊 Generating complex Mermaid chart with multiple arrows...")
            try:
                mermaid_code = generate_mermaid_chart(slide_data, topic)
                mermaid_image_path = os.path.join(charts_dir, f"chart_{idx:02d}.png")
                render_mermaid_to_image(mermaid_code, mermaid_image_path)
            except Exception as e:
                print(f"  ⚠️  Skipping chart for this slide: {e}")
                mermaid_image_path = None
        else:
            print(f"  ℹ️  Title slide - no chart needed")
        
        print(f"  📝 Generating narration script...")
        narration_text = generate_narration_script(slide_data)
        print(f"  Script: {narration_text[:100]}...")
        
        # Generate speech
        audio_filename = os.path.join(audio_dir, f"slide_{idx:02d}.mp3")
        duration = generate_speech_elevenlabs(narration_text, audio_filename)
        
        # Create slide image
        slide_image_path = os.path.join(slides_dir, f"slide_{idx:02d}.png")
        create_slide_image(slide_data, logo_path, slide_image_path, mermaid_image_path)
        
        # Create video clip
        image_clip = ImageClip(slide_image_path).set_duration(duration)
        audio_clip = AudioFileClip(audio_filename)
        video_clip = image_clip.set_audio(audio_clip)
        
        video_clips.append(video_clip)
        print(f"  ✅ Video clip created (Duration: {duration:.1f}s)")
    
    print(f"\n🎬 Combining all slides into final video...")
    final_video = concatenate_videoclips(video_clips, method="compose")
    
    print(f"💾 Rendering final video...")
    final_video.write_videofile(
        output_filename,
        fps=24,
        codec='libx264',
        audio_codec='aac',
        temp_audiofile='temp-audio.m4a',
        remove_temp=True
    )
    
    final_video.close()
    for clip in video_clips:
        clip.close()
    
    print(f"\n✨ Video presentation saved as: {output_filename}")
    print(f" Audio files: {audio_dir}/")
    print(f"  Slide images: {slides_dir}/")
    print(f" Mermaid charts: {charts_dir}/")
    print(f" Total slides: {len(slides_data)}")
    print(f" Learning data saved to: {LEARNING_FILE}")
    
    return output_filename

if __name__ == "__main__":
    generate_video_presentation("Python Tuples and Their Applications")