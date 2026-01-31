import os
import subprocess
import sys
import shutil

def build():
    # Detect OS
    is_win = sys.platform.startswith('win')
    
    # Clean dist/build
    if os.path.exists('dist'): shutil.rmtree('dist')
    if os.path.exists('build'): shutil.rmtree('build')
    
    # Command
    # --onefile: single exe
    # --windowed: no console
    # --add-data: products.json, config.json (Note: these need to be handled in code to look at sys._MEIPASS if bundled, OR keep them external)
    # Actually, for data files that need to be edited (products/history), they should NOT be bundled inside the ONEFILE EXE because they will be read-only or temporary.
    # They should be external.
    # So we do NOT add them to --add-data, but we ensure the code looks for them in the executable directory.
    
    # In logic.py/history.py, we used:
    # BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # If frozen, sys.executable is the path.
    # We need to ensure logic.py handles "frozen" state to find external files.
    
    # But for now, let's just build.
    
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name=DeliveryOrderGen',
        '--windowed',
        '--onefile',
        'src/main.py'
    ]
    
    print(f"Building with command: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)
    
    print("Build complete. Output in dist/")

if __name__ == "__main__":
    build()
