"""Clear and regenerate comtypes cache to prevent import errors."""
import shutil
from pathlib import Path
import comtypes.client

def clear_cache():
    """Remove all comtypes generated files and regenerate UIAutomation."""
    # Find comtypes gen directory
    gen_dir = Path(comtypes.client.__file__).parent / "gen"
    
    if gen_dir.exists():
        print(f"Clearing comtypes cache at: {gen_dir}")
        # Remove all files except __pycache__ and __init__.py
        for item in gen_dir.iterdir():
            if item.name not in ("__pycache__", "__init__.py"):
                if item.is_file():
                    item.unlink()
                    print(f"  Removed: {item.name}")
    
    # Regenerate UIAutomation type library
    print("\nRegenerating UIAutomation type library...")
    comtypes.client.GetModule('UIAutomationCore.dll')
    print("Done! Comtypes cache cleared and regenerated.")

if __name__ == "__main__":
    clear_cache()
