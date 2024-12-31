import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import HexColor, Color
import io
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from reportlab.graphics.shapes import Drawing, Group, Path, String

# Function to change color to light gray
def change_color_to_gray(drawing):
    light_gray = Color(0.7294117647, 0.6862745098, 0.6274509804)  # Light gray color
    for element in drawing.contents:
        if hasattr(element, 'fillColor') and element.fillColor:
            element.fillColor = light_gray
        if isinstance(element, Group):
            change_color_to_gray(element)
        elif isinstance(element, Path) or isinstance(element, String):
            element.fillColor = light_gray

# Function to create a new PDF page with the participant's name
def create_certificate(template_path, name, transparency_values):
    # Convert the name to uppercase
    name = name.upper()
    
    # Read the existing PDF template
    reader = PdfReader(template_path)
    writer = PdfWriter()

    # Create a canvas to draw the name on
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    # Register and set the custom font
    pdfmetrics.registerFont(TTFont('<Add Font Name>', r'<PATH_TO_TTF_FONT>'))

    # Calculate the width of the text and adjust font size if necessary
    max_width = letter[0] - 200  # Max width with some padding
    font_size = 45.3  # Starting font size
    while can.stringWidth(name, "<Add Font Name>", font_size) > max_width and font_size > 10:
        font_size -= 15

    can.setFont("<Add Font Name>", font_size)
    
    # Set the color
    fill_color = HexColor("#9a6206")
    can.setFillColor(fill_color)
    can.setStrokeColor(fill_color)

    # Calculate the width of the text and center-align it
    text_width = can.stringWidth(name, "<Add Font Name>", font_size)
    x = (letter[0] - text_width) / 2 - 7.9 # Centered x position
    y = 322  # y position

    # Simulating bold text by drawing multiple times with small offsets
    offset = 0.5  # Adjust the offset for the desired bold effect
    can.drawString(x, y, name)  # Original position
    can.drawString(x + offset, y, name)  # Slightly right
    can.drawString(x - offset, y, name)  # Slightly left
    can.drawString(x, y + offset, name)  # Slightly up
    can.drawString(x, y - offset, name)  # Slightly down

    # Load and draw SVG images
    svg_paths = [
        r'<PATH_TO_SVG_FILES>',
    
        
    ]
    
    # Adjusted SVG positions
    svg_positions = [
        (99, 695), (216, 695), (333, 695),
        (450, 695),
    ]

    # Desired size in points (14.22 mm converted to points)
    desired_width_in_points = 14.22 * (72 / 25.4)
    desired_height_in_points = 14.22 * (72 / 25.4)

    for svg_path, (svg_x, svg_y), transparency in zip(svg_paths, svg_positions, transparency_values):
        drawing = svg2rlg(svg_path)
        original_width = drawing.width
        original_height = drawing.height
        
        # Calculate scaling factors
        scale_x = desired_width_in_points / original_width
        scale_y = desired_height_in_points / original_height
        
        # Apply scaling
        drawing.scale(scale_x, scale_y)
        
        # Change color to light gray if transparency is 0
        if transparency == 0:
            change_color_to_gray(drawing)

        renderPDF.draw(drawing, can, svg_x, svg_y)
    
    can.save()

    # Move to the beginning of the BytesIO buffer
    packet.seek(0)
    new_pdf = PdfReader(packet)

    # Add the "watermark" (which is the new PDF) on the existing page
    page = reader.pages[0]
    page.merge_page(new_pdf.pages[0])
    
    return page

# Path to the PDF template
template_path = r'<PATH_TO_PDF_TEMPLATE>'

# Path to the Excel file with participant names and transparency values
excel_file = r'<PATH_TO_EXCEL_FILE>'

# Path to the output file
output_file = r'<PATH_TO_OUTPUT_DIRECTORY>/All Certificates.pdf'

# Read the Excel file
df = pd.read_excel(excel_file)

# Create a PdfWriter object for the combined PDF
combined_writer = PdfWriter()

# Loop through each participant's name and add their certificate to the combined PDF
for index, row in df.iterrows():
    name = row['Name']
    transparency_values = row[1:].tolist()  # Get transparency values from the row
    certificate_page = create_certificate(template_path, name, transparency_values)
    combined_writer.add_page(certificate_page)

# Write the combined PDF to a file
with open(output_file, "wb") as outputStream:
    combined_writer.write(outputStream)

print("Combined certificates created successfully!")
