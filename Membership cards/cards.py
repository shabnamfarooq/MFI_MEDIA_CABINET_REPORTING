from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os

# DPI setting for the image
DPI = 300

# Convert inches to pixels based on DPI
def inches_to_pixels(inches, dpi):
    return int(inches * dpi)

# Function to generate membership cards
def generate_membership_cards(csv_path, template_path, font_path, output_dir):
    # Load the member data from a CSV file with a specified encoding
    members_df = pd.read_csv(csv_path, encoding='ISO-8859-1')

    # Load the base template image
    template = Image.open(template_path)

    # Define the bold font and size (Ensure the bold font file is accessible)
    font = ImageFont.truetype(font_path, 36)

    # Ensure the cards directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Text dimensions and positions in inches
    text_width_in = 2.55
    text_height_in = 0.13
    name_x_in = 0.21
    name_y_in = 1.09
    member_id_y_in = 1.25  # Adjusted for spacing between lines

    # Convert dimensions to pixels
    text_width_px = inches_to_pixels(text_width_in, DPI)
    text_height_px = inches_to_pixels(text_height_in, DPI)
    name_x_px = inches_to_pixels(name_x_in, DPI)
    name_y_px = inches_to_pixels(name_y_in, DPI)
    member_id_y_px = inches_to_pixels(member_id_y_in, DPI)

    # Function to draw centered text within a given width
    def draw_centered_text(draw, text, font, x, y, width):
        text_bbox = draw.textbbox((0, 0), text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x_position = x + (width - text_width) / 2
        draw.text((x_position, y), text, font=font, fill='White')

    # Iterate over each member and generate a card
    for index, row in members_df.iterrows():
        card = template.copy()
        draw = ImageDraw.Draw(card)

        # Prepare the text to be drawn
        name_text = f"Name: {row['Name']}"
        member_id_text = f"MemberID: {row['MemberID']}"

        # Draw the text on the card
        draw_centered_text(draw, name_text, font, name_x_px, name_y_px, text_width_px)
        draw_centered_text(draw, member_id_text, font, name_x_px, member_id_y_px, text_width_px)

        # Save the card
        card.save(os.path.join(output_dir, f'membership_card_{row["MemberID"]}.png'))

    print("Membership cards have been generated successfully!")

# Usage example (ensure paths are correct)
csv_path = 'members.csv'
template_path = 'membership_card_template.png'
font_path = 'calibrib.ttf'
output_dir = ''


generate_membership_cards(csv_path, template_path, font_path, output_dir)
