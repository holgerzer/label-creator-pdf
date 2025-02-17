# label-creator-pdf

 A tool to create multiple labels on a PDF page.

# Install packages

`pip install pandas reportlab openpyxl`

# Create executable

`pyinstaller --onefile --add-data "stag-sans-thin.ttf;." --add-data "stag-sans-book.ttf;." --add-data "stag-sans-light-italic.ttf;." --add-data "stag-sans-medium.ttf;." label_print.py`