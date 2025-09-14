#!/usr/bin/env python3
"""
Build script for Email Automation Tool
Creates a standalone executable from the GUI application using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

# Configuration
APP_NAME = "EmailAutomation"
MAIN_SCRIPT = "App/gui.py"
OUTPUT_DIR = "App/output"
DIST_DIR = "dist"
BUILD_DIR = "build"


def check_pyinstaller():
    """Check if PyInstaller is installed, install if not"""
    try:
        import PyInstaller

        print("✓ PyInstaller is already installed")
        return True
    except ImportError:
        print("PyInstaller not found. Installing...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "pyinstaller"],
                check=True,
                capture_output=True,
                text=True,
            )
            print("✓ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install PyInstaller: {e}")
            return False


def install_dependencies():
    """Install all required dependencies from requirements.txt"""
    if os.path.exists("requirements.txt"):
        print("Installing dependencies from requirements.txt...")
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
                check=True,
                capture_output=True,
                text=True,
            )
            print("✓ Dependencies installed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install dependencies: {e}")
            return False
    else:
        print("⚠ requirements.txt not found, skipping dependency installation")
        return True


def clean_build_dirs():
    """Clean previous build directories"""
    dirs_to_clean = [BUILD_DIR, DIST_DIR]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"Cleaning {dir_name}...")
            shutil.rmtree(dir_name)


def build_executable():
    """Build the executable using PyInstaller"""
    if not os.path.exists(MAIN_SCRIPT):
        print(f"✗ Main script not found: {MAIN_SCRIPT}")
        return False

    print(f"Building executable from {MAIN_SCRIPT}...")

    # PyInstaller command arguments
    pyinstaller_args = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",  # Create a single executable file
        "--windowed",  # Hide console window (GUI app)
        "--name",
        APP_NAME,
        "--distpath",
        OUTPUT_DIR,
        "--workpath",
        BUILD_DIR,
        "--specpath",
        ".",
        # Add any additional data files if needed
        # "--add-data", "assets;assets",  # Uncomment if you have assets
        MAIN_SCRIPT,
    ]

    try:
        result = subprocess.run(
            pyinstaller_args, check=True, capture_output=True, text=True
        )
        print("✓ Executable built successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Build failed: {e}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return False


def create_output_dir():
    """Create output directory if it doesn't exist"""
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    print(f"✓ Output directory created: {OUTPUT_DIR}")


def print_build_info():
    """Print information about the build"""
    exe_path = os.path.join(OUTPUT_DIR, f"{APP_NAME}.exe")
    if os.path.exists(exe_path):
        file_size = os.path.getsize(exe_path)
        file_size_mb = file_size / (1024 * 1024)
        print(f"\n{'='*50}")
        print(f"BUILD SUCCESSFUL!")
        print(f"{'='*50}")
        print(f"Executable: {exe_path}")
        print(f"Size: {file_size_mb:.2f} MB")
        print(f"{'='*50}")
        print("\nYou can now run the application by double-clicking the .exe file")
        print("or by running it from the command line.")
    else:
        print("\n✗ Build completed but executable not found in expected location")


def main():
    """Main build process"""
    print("Email Automation Tool - Build Script")
    print("=" * 40)

    # Step 1: Check and install PyInstaller
    if not check_pyinstaller():
        sys.exit(1)

    # Step 2: Install dependencies
    if not install_dependencies():
        print("⚠ Continuing build despite dependency installation issues...")

    # Step 3: Clean previous builds
    clean_build_dirs()

    # Step 4: Create output directory
    create_output_dir()

    # Step 5: Build executable
    if not build_executable():
        sys.exit(1)

    # Step 6: Print build information
    print_build_info()

    # Step 7: Clean up temporary files
    if os.path.exists(BUILD_DIR):
        shutil.rmtree(BUILD_DIR)
        print("✓ Cleaned up temporary build files")


if __name__ == "__main__":
    main()
