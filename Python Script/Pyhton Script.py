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
    pdfmetrics.registerFont(TTFont('Playfair_Display', r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\Font\Playfair_Display\static\PlayfairDisplay-Medium.ttf'))

    # Calculate the width of the text and adjust font size smoothly
    max_width = letter[0] - 200  # Max text width with margin
    max_font_size = 45
    min_font_size = 34
    font_size = max_font_size

    # Shrink font only if name is too wide
    while can.stringWidth(name, "Playfair_Display", font_size) > max_width and font_size > min_font_size:
        font_size -= 1

    can.setFont("Playfair_Display", font_size)

    # Set the font color
    fill_color = HexColor("#d49c3d") # Gold color
    can.setFillColor(fill_color)
    can.setStrokeColor(fill_color)

    # Calculate the width of the text and center-align it
    text_width = can.stringWidth(name, "Playfair_Display", font_size)
    x = (letter[0] - text_width) / 2 - 9.5  # horizontal center with slight left margin
    y = 287  # vertical position of the name

    # Draw the name on the certificate
    can.drawString(x, y, name)

    # Log the font size used
    print(f"Name: {name} → Font Size: {font_size}")

    # Load and draw SVG images
    svg_paths = [
        r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\All WCA Events Logo\333.svg',
        r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\All WCA Events Logo\444.svg',
        r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\All WCA Events Logo\333oh.svg',
        r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\All WCA Events Logo\minx.svg',
    ]
    
    svg_positions = [
        (99, 695), (216, 695), (333, 695),
        (450, 695),
    ]

    # Size conversion from mm to points
    desired_width_in_points = 14.22 * (72 / 25.4)
    desired_height_in_points = 14.22 * (72 / 25.4)

    for svg_path, (svg_x, svg_y), transparency in zip(svg_paths, svg_positions, transparency_values):
        drawing = svg2rlg(svg_path)
        if drawing is None:
            print(f"⚠️ Failed to load SVG: {svg_path}")
            continue

        scale_x = desired_width_in_points / drawing.width
        scale_y = desired_height_in_points / drawing.height
        drawing.scale(scale_x, scale_y)

        if transparency == 0:
            change_color_to_gray(drawing)

        renderPDF.draw(drawing, can, svg_x, svg_y)
    
    can.save()

    # Merge the new canvas with the template PDF
    packet.seek(0)
    new_pdf = PdfReader(packet)
    page = reader.pages[0]
    page.merge_page(new_pdf.pages[0])
    
    return page

# File paths
template_path = r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\Template.pdf'
excel_file = r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\names.xlsx'
output_file = r'D:\Kushaan all work\Activities\Cubing Competitions\Delhi Cube Autumn Open\Output\All Certificates.pdf'

# Read the Excel file
df = pd.read_excel(excel_file)

# Create a PdfWriter object for the final combined PDF
combined_writer = PdfWriter()

# Generate and collect each certificate
for index, row in df.iterrows():
    name = row['Name']
    transparency_values = row[1:].tolist()
    certificate_page = create_certificate(template_path, name, transparency_values)
    combined_writer.add_page(certificate_page)

# Save the final PDF
with open(output_file, "wb") as outputStream:
    combined_writer.write(outputStream)

print("\n✅ Combined certificates created successfully!")
