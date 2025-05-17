# setup.py
import os
import subprocess
import sys
import platform

def create_virtual_env():
    """Create and setup a virtual environment for the document metadata extractor"""
    print("Setting up virtual environment for Document Metadata Extractor...")
    
    # Define directories and paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    venv_dir = os.path.join(base_dir, "venv")
    requirements_path = os.path.join(base_dir, "requirements.txt")
    
    # Check if virtual environment exists
    if os.path.exists(venv_dir):
        print("Virtual environment already exists.")
    else:
        print("Creating virtual environment...")
        subprocess.check_call([sys.executable, "-m", "venv", venv_dir])
    
    # Determine the pip and python paths in the virtual environment
    if platform.system() == "Windows":
        python_path = os.path.join(venv_dir, "Scripts", "python.exe")
        pip_path = os.path.join(venv_dir, "Scripts", "pip.exe")
    else:
        python_path = os.path.join(venv_dir, "bin", "python")
        pip_path = os.path.join(venv_dir, "bin", "pip")
    
    # Install requirements
    print("Installing required packages...")
    subprocess.check_call([pip_path, "install", "-r", requirements_path])
    
    print("\nSetup complete! You can now run the application with:")
    
    if platform.system() == "Windows":
        print(f"{python_path} main.py")
    else:
        print(f"{python_path} main.py")
    
    return python_path

if __name__ == "__main__":
    create_virtual_env()