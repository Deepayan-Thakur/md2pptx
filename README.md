# md2pptx 🚀

**AI-Powered Markdown to Professional PPTX Generator**  
*Built for the Accenture EZ Hackathon 2026*

`md2pptx` transforms raw Markdown files into professional, executive-ready PowerPoint presentations. Unlike standard converters, it uses advanced Reasoning LLMs (DeepSeek-R1) to synthesize high-value content and dynamic AI-generated visual assets via Hugging Face.

## ✨ Key Features

*   **🧠 Deep Reasoning Synthesis**: Utilizes `DeepSeek-R1-Distill-Qwen-32B` via Hugging Face to process up to 150,000 characters and generate insightful summaries.
*   **📊 Recursive Data Precision**: Deterministically extracts numerical data from Markdown tables to build accurate, professional graphs using Matplotlib.
*   **🎨 Structural Aesthetic Boxes**: Replaces boring bullet lists with structural UI "Creative Boxes" featuring dynamic accent stripes.
*   **🖼️ AI-Generated Assets**: Dynamically generates slide-specific corporate illustrations using the **FLUX.1-schnell** image model.
*   **⚡ High Performance**: Fast PPTX construction with local image caching to respect API rate limits.
*   **🏗️ Flexible Providers**: Support for both **Gemini 2.0 Flash** and **Hugging Face** Inference API.

## 🛠️ Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/md2pptx.git
   cd md2pptx
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Configure environment variables in a `.env` file:
   ```env
   HUGGINGFACE_API_KEY=your_key_here
   GEMINI_API_KEY=your_key_here
   ```

## 🚀 Usage

Run the generator from the project root:

```bash
python main.py path/to/content.md
```

### Advanced Options

*   `--slides N`: Target specific slide count (Default: 13).
*   `--provider [huggingface|gemini]`: Choose your AI backend (Default: huggingface).

## 📁 Project Structure

```text
├── main.py              # Application Entry Point
├── md2pptx/
│   ├── src/             # Core Modules (Parser, Planner, Builder, ImageGen)
│   ├── assets/          # Brand Identity Template & Static Assets
│   └── outputs/         # Generated Presentations
└── test_cases/          # Sample Markdown Datasets
```
## Setup Instructions
1. Clone repository
2. Create virtual environment: `python -m venv .venv`
3. Activate: `.\.venv\Scripts\Activate.ps1`
4. Install dependencies: `pip install -r requirements.txt`
5. Set environment variables (create `.env`):
   - `OPENROUTER_API_KEY=your_key_here`
   - `GEMINI_API_KEY=your_key_here`
6. Run: `python main.py test_cases/example.md`

## 📄 License

This project was developed for the Accenture EZ Hackathon 2026.
