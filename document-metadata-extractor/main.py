# main.py
import os
import sys
import importlib.util
import subprocess

def check_venv():
    """Check if running in virtual environment and activate if needed"""
    # Check if already in a virtual environment
    if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        return True
    
    # If not in virtual environment, try to activate it
    base_dir = os.path.dirname(os.path.abspath(__file__))
    venv_dir = os.path.join(base_dir, "venv")
    
    if not os.path.exists(venv_dir):
        print("Virtual environment not found. Running setup.py...")
        # Run setup.py to create and configure virtual environment
        python_path = None
        try:
            spec = importlib.util.spec_from_file_location("setup", os.path.join(base_dir, "setup.py"))
            setup = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(setup)
            python_path = setup.create_virtual_env()
        except Exception as e:
            print(f"Error setting up virtual environment: {e}")
            return False
        
        # Relaunch the script with the virtual environment's Python
        if python_path:
            script_path = os.path.abspath(__file__)
            subprocess.call([python_path, script_path])
            sys.exit(0)
            
        return False
    
    # If virtual environment exists but not activated, print instructions
    print("Please run this script from the virtual environment:")
    if os.name == 'nt':  # Windows
        print(f".\\venv\\Scripts\\python.exe main.py")
    else:  # Unix/Linux/Mac
        print(f"source venv/bin/activate && python main.py")
        
    return False

def run_application():
    """Run the document metadata extractor application"""
    # Import the document metadata extractor module
    from document_metadata_extractor import run_application
    run_application()

if __name__ == "__main__":
    if check_venv():
        run_application()