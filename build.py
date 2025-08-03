import PyInstaller.__main__
import os
import sys

def build_app():
    # Detect platform for correct separator
    separator = ':' if os.name == 'posix' else ';'
    
    # Common arguments
    args = [
        'main.py',
        '--onefile',  # Single executable
        '--windowed',  # No console window
        '--name=GamingLoungeManager',
        f'--add-data=config.json{separator}.',  # Include config file
    ]
    
    # Add icon if exists
    if os.path.exists('icon.ico'):
        args.append('--icon=icon.ico')
    
    PyInstaller.__main__.run(args)

if __name__ == "__main__":
    build_app()
