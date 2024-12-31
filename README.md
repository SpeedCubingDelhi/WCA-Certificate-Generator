WCA-Certificate-Generator is a Python script designed to automate the creation of WCA competition certificates of participation. This script customizes certificates by adding participant names, event logos, and dynamic font styles based on a provided template.

Requirements:
A basic certificate template (PDF or image format).
An Excel sheet with participant names (downloaded from WCA).
A .ttf font file of your choice for name styling.
SVG images of the event logos to place on the certificate.
Outcome:
All certificates will be generated with the participant names automatically positioned.
Users can customize the x and y coordinates of the names to suit their specific template.
Event logos are dynamically styled:
Logos for events in which the participant is competing are highlighted.
Logos for non-participated events are displayed with reduced transparency.
For example: If a person participates in 3x3, 2x2, and 4x4 but not 5x5, the 5x5 logo will appear less prominent on the certificate.
This repository is a complete solution for generating professional certificates tailored for WCA competitions, ensuring accuracy and customization for every participant.
