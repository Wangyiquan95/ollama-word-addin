from PIL import Image, ImageDraw, ImageFont
import os

def create_icon(size, filename):
    # Create a new image with a blue background
    img = Image.new('RGBA', (size, size), (0, 120, 212, 255))
    draw = ImageDraw.Draw(img)
    
    # Try to use a system font, fallback to default
    try:
        font_size = max(8, size // 2)
        font = ImageFont.truetype("/System/Library/Fonts/Arial.ttf", font_size)
    except:
        font = ImageFont.load_default()
    
    # Draw 'O' in the center
    text = "O"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    x = (size - text_width) // 2
    y = (size - text_height) // 2
    
    draw.text((x, y), text, fill='white', font=font)
    
    # Save the image
    img.save(filename, 'PNG')
    print(f"Created {filename}")

# Create all required icon sizes
create_icon(16, 'assets/icon-16.png')
create_icon(32, 'assets/icon-32.png')
create_icon(64, 'assets/icon-64.png')
create_icon(80, 'assets/icon-80.png')
create_icon(90, 'assets/logo-filled.png')

print("All icons created successfully!")
