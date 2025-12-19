# Guide: Generating PowerPoint Presentations (PPTX)

This guide will walk you through the steps to generate a PowerPoint presentation (.pptx file) from the Escalation Details Report data.

## Prerequisites

- Python 3.7 or higher installed on your system
- Access to the project directory
- Terminal/Command line access

## Step-by-Step Instructions

### Step 1: Navigate to the Project Directory

Open your terminal and navigate to the project directory:

```bash
cd /Users/kylespears/Cursor/Escalation-Details-Report
```

(Adjust the path if your project is located elsewhere)

### Step 2: Create a Virtual Environment (First Time Only)

If you haven't already set up a virtual environment, create one:

```bash
python3 -m venv venv
```

This creates a new directory called `venv` in your project folder. This virtual environment will keep all the required packages isolated from your system Python installation.

### Step 3: Activate the Virtual Environment

**Important:** You must activate the virtual environment every time you open a new terminal session.

**On macOS/Linux:**
```bash
source venv/bin/activate
```

**On Windows:**
```bash
venv\Scripts\activate
```

When activated, you should see `(venv)` at the beginning of your terminal prompt, like this:
```
(venv) kylespears@Mac Escalation-Details-Report %
```

### Step 4: Install Required Dependencies

Install all required packages from the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

This will install:
- `python-pptx` (for creating PowerPoint presentations)
- `playwright` (for rendering charts)
- `Pillow` (for image processing)
- And their dependencies

**Note:** If you only need to install `python-pptx` specifically, you can run:
```bash
pip install python-pptx
```

### Step 5: Install Playwright Browsers (Required for Charts)

**Important:** After installing Playwright, you must also install the browser binaries. This is required for chart rendering to work:

```bash
playwright install chromium
```

This will download Chromium (~160 MB) which is used to render the HTML charts as images. Without this step, your charts will appear blank in the generated PowerPoint.

**Note:** This is a one-time setup. The browsers are installed globally and will be reused for future runs.

### Step 6: Verify Installation

Verify that `python-pptx` is installed correctly:

```bash
pip list | grep pptx
```

You should see output like:
```
python-pptx       1.0.2
```

### Step 7: Run the Presentation Generator

Run the script to generate your PowerPoint presentation:

```bash
python3 generate_presentation.py
```

#### Available Options

The script supports several command-line options:

- **`--client-name CLIENT_NAME`**: Override the default client name in the report
  ```bash
  python3 generate_presentation.py --client-name "Acme Corporation"
  ```

- **`--no-threat-landscape`**: Exclude threat landscape slides from the presentation
  ```bash
  python3 generate_presentation.py --no-threat-landscape
  ```

- **`--output-dir OUTPUT_DIR`**: Specify a custom output directory (default: `output`)
  ```bash
  python3 generate_presentation.py --output-dir my_reports
  ```

- **`--keep-charts`**: Keep temporary chart images after generation (default: charts are deleted)
  ```bash
  python3 generate_presentation.py --keep-charts
  ```

- **`--help`**: Display all available options
  ```bash
  python3 generate_presentation.py --help
  ```

### Step 8: Locate Your Generated PPTX File

After the script completes successfully, your PowerPoint presentation will be saved in the `output/` directory (or your specified output directory).

The file will be named something like:
- `escalation_details_report_YYYY-MM-DD_HH-MM-SS.pptx`

You can open this file with Microsoft PowerPoint, Google Slides, or any compatible presentation software.

## Complete Example Workflow

Here's a complete example of the entire process from start to finish:

```bash
# 1. Navigate to project directory
cd /Users/kylespears/Cursor/Escalation-Details-Report

# 2. Activate virtual environment
source venv/bin/activate

# 3. (First time only) Install dependencies
pip install -r requirements.txt

# 4. (First time only) Install Playwright browsers for chart rendering
playwright install chromium

# 5. Generate the presentation
python3 generate_presentation.py --client-name "Example Corp"

# 6. Find your file in the output/ directory
ls -lh output/
```

## Troubleshooting

### Issue: "ModuleNotFoundError: No module named 'pptx'"

**Solution:** Make sure you've activated the virtual environment and installed the dependencies:
```bash
source venv/bin/activate
pip install -r requirements.txt
```

### Issue: Charts are blank or missing in the PowerPoint

**Solution:** This usually means Playwright browsers aren't installed. Install Chromium:
```bash
source venv/bin/activate
playwright install chromium
```

Then regenerate your presentation. The charts are rendered as images using Playwright, so the browser binaries are required.

### Issue: "ImportError: cannot import name 'RGBColor' from 'pptx.util'"

**Solution:** This has been fixed in the code. If you encounter this, make sure you have the latest version of the script. The correct import is:
```python
from pptx.dml.color import RGBColor
```

### Issue: Virtual environment not activating

**Solution:** 
- Make sure you're in the project directory
- Check that the `venv` directory exists
- Try recreating the virtual environment:
  ```bash
  rm -rf venv
  python3 -m venv venv
  source venv/bin/activate
  ```

### Issue: Permission errors when installing packages

**Solution:** Make sure you're using the virtual environment (you should see `(venv)` in your prompt). Never use `sudo` with pip when working in a virtual environment.

### Issue: Script runs but no output file is created

**Solution:**
- Check the terminal output for error messages
- Verify that the `output/` directory exists and is writable
- Check that you have the required data files (e.g., `Report Extract.xlsx`)

## Quick Reference Commands

| Task | Command |
|------|---------|
| Task | Command |
|------|---------|
| Activate venv | `source venv/bin/activate` |
| Deactivate venv | `deactivate` |
| Install dependencies | `pip install -r requirements.txt` |
| Install Playwright browsers | `playwright install chromium` |
| Check installed packages | `pip list` |
| Run script | `python3 generate_presentation.py` |
| See script options | `python3 generate_presentation.py --help` |

## Important Notes

1. **Always activate the virtual environment** before running the script. If you forget, you'll get import errors.

2. **The virtual environment is project-specific**. Each time you open a new terminal, you need to activate it again.

3. **Don't commit the `venv` directory** to version control. It's already in `.gitignore`.

4. **If you update dependencies**, run `pip install -r requirements.txt` again to ensure you have the latest versions.

## Getting Help

If you encounter issues not covered in this guide:

1. Check the terminal output for specific error messages
2. Verify all prerequisites are met
3. Ensure the virtual environment is activated
4. Try reinstalling dependencies: `pip install --upgrade -r requirements.txt`

## Next Steps

Once you've generated your PPTX file:
- Open it in Microsoft PowerPoint or compatible software
- Review the slides for accuracy
- Customize any content as needed
- Share with your team or clients

---

**Last Updated:** This guide covers the setup process as of the current project state. If you encounter any issues, please refer to the troubleshooting section above.

