# label-creator-pdf

 A tool to create multiple labels on a PDF page.

# Install packages

` pip install -r requirements.txt`

# Create executable

## Start clean virtual environment

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```
## Build .exe-file

`pyinstaller --onefile --noconsole --clean --strip --paths=src --hidden-import=gui --hidden-import=label_creator src/main.py`