import os
import shutil
import subprocess
from PIL import Image

def create_icons(source_path):
    if not os.path.exists(source_path):
        print(f"Error: {source_path} not found.")
        return

    img = Image.open(source_path)
    
    # Ensure standard RGB, handle transparency if png, but here it's jpg so no alpha usually.
    # If user provided JPG, we might want to mask it to circle or rounded rect?
    # For now, just use as is. 
    # Resize to max needed
    img = img.convert("RGBA")
    
    # 1. Create .ico (Windows)
    # Windows likes 256, 128, 64, 48, 32, 16
    ico_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    img.save("resources/app_icon.ico", format='ICO', sizes=ico_sizes)
    print("Created resources/app_icon.ico")

    # 2. Create .icns (macOS) using iconutil
    # Needs a .iconset directory with specific names
    iconset_dir = "resources/app_icon.iconset"
    if os.path.exists(iconset_dir):
        shutil.rmtree(iconset_dir)
    os.makedirs(iconset_dir)

    # Specs:
    # icon_16x16.png
    # icon_16x16@2x.png (32x32)
    # icon_32x32.png
    # icon_32x32@2x.png (64x64)
    # icon_128x128.png
    # icon_128x128@2x.png (256x256)
    # icon_256x256.png
    # icon_256x256@2x.png (512x512)
    # icon_512x512.png
    # icon_512x512@2x.png (1024x1024)
    
    specs = [
        ("icon_16x16.png", 16),
        ("icon_16x16@2x.png", 32),
        ("icon_32x32.png", 32),
        ("icon_32x32@2x.png", 64),
        ("icon_128x128.png", 128),
        ("icon_128x128@2x.png", 256),
        ("icon_256x256.png", 256),
        ("icon_256x256@2x.png", 512),
        ("icon_512x512.png", 512),
        ("icon_512x512@2x.png", 1024),
    ]

    for filename, size in specs:
        resized = img.resize((size, size), Image.Resampling.LANCZOS)
        resized.save(os.path.join(iconset_dir, filename))

    # Run iconutil
    try:
        subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", "resources/app_icon.icns"], check=True)
        print("Created resources/app_icon.icns")
    except subprocess.CalledProcessError as e:
        print(f"Error creating icns: {e}")
    except FileNotFoundError:
        print("iconutil not found (not on macOS?)")

    # Cleanup iconset
    # shutil.rmtree(iconset_dir)

if __name__ == "__main__":
    create_icons("resources/original_icon.jpg")
