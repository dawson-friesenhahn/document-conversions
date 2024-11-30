
All of these instructions were written for Windows, using git bash as the terminal. I'm running Python 3.12.6

# Setup

1. Clone the repository, set as working directory in terminal
2. Set up and activate a python virtual environment

```
python -m venv .venv
. .venv/Scripts/activate
```

3. Install dependencies: `pip install -r requirements.txt`
4. Modify `port` argument in [main.py](main.py), if necessary
5. Run the conversion server: `python main.py`

# Usage

The server accepts POST requests containing ppt and pptx files at the `/pptx-to-pdf` endpoint. For example usage, see [test.py](test.py)




