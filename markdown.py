import re
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from markdown2 import markdown
from bs4 import BeautifulSoup

def parse_markdown_to_docx(md_file, docx_file):
    # Read the markdown file
    with open(md_file, 'r', encoding='utf-8') as file:
        md_content = file.read()

    # Convert markdown to HTML
    html_content = markdown(md_content)

    # Parse HTML content
    soup = BeautifulSoup(html_content, 'html.parser')

    # Create a new Word document
    doc = Document()

    # Add a title
    doc.add_heading('Markdown to Word Conversion', 0)

    # Function to add paragraph with optional bold and italic text
    def add_paragraph(text, style=None, bold=False, italic=False):
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(text)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if style:
            paragraph.style = style
        return paragraph

    # Parse HTML elements and add them to the Word document
    for element in soup.children:
        if element.name == 'p':
            add_paragraph(element.text)
        elif element.name == 'h1':
            add_paragraph(element.text, 'Heading 1')
        elif element.name == 'h2':
            add_paragraph(element.text, 'Heading 2')
        elif element.name == 'h3':
            add_paragraph(element.text, 'Heading 3')
        elif element.name == 'ul':
            for li in element.find_all('li'):
                add_paragraph('• ' + li.text)
        elif element.name == 'blockquote':
            quote = add_paragraph(element.text, style='Quote')
            quote.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif element.name == 'strong':
            add_paragraph(element.text, bold=True)
        elif element.name == 'em':
            add_paragraph(element.text, italic=True)
        elif element.name == 'code':
            add_paragraph(element.text, style='Code')
        else:
            add_paragraph(element.text)

    # Save the document
    doc.save(docx_file)

# Convert the given Markdown file to Word document
md_file_path = 'path/to/your/markdown/file.md'
docx_file_path = 'path/to/your/output/file.docx'
parse_markdown_to_docx(md_file_path, docx_file_path)
