# AI Video Presentation Generator

Generate complete video presentations from a single topic using AI—automated slides, narration, diagrams, and video assembly.

[Watch Demo Video](https://www.youtube.com/watch?v=oJbCWTZbRz4)

---

## Features

- **Auto-Generated Content**: GPT-4 creates slide structure and bullet points
- **Complex Mermaid Diagrams**: Dynamic flowcharts with multiple relationships
- **Natural Narration**: AI-powered scripts with ElevenLabs voice synthesis
- **Video Assembly**: Combines slides, audio, and timing into MP4
- **Adaptive Learning**: Learns from successful diagram patterns to avoid syntax errors
- **Logo Branding**: Optional company or personal branding on slides

---

## Prerequisites

- Python 3.8+
- OpenAI API key
- ElevenLabs API key and Voice ID

---

## Installation

### Clone Repository

```bash
git clone https://github.com/savir2010/video-gen.git
cd video-gen
```

### Create Virtual Environment

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Configuration

Create a `.env` file in the project root:

```env
OPENAI_KEY=sk-...
ELEVENLABS_API_KEY=...
ELEVENLABS_VOICE_ID=...
```

**Optional**: Add `logo.png` in the project root for slide branding.

---

## Usage

Run the generator:

```bash
python main.py
```

**Customize the topic** by editing the last line in `main.py`:

```python
generate_video_presentation("Your Topic Here")
```

---

## Output Structure

```
project/
├── audio/          # Generated narration files (.mp3)
├── slides/         # Slide images (.png)
├── charts/         # Mermaid diagram images (.png)
├── output.mp4      # Final video presentation
└── mermaid_patterns.pkl  # Learning data for diagram generation
```

