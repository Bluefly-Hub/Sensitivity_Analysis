"""
Auto-updater for Cerberus Sensitivity Analysis
Checks GitHub releases and handles automatic updates
"""

import urllib.request
import json
import os
import sys
import subprocess
import shutil
import tkinter as tk
from tkinter import messagebox
from version import __version__


class AutoUpdater:
    """Handles checking for updates and automatic application updates"""
    
    def __init__(self, repo_owner="Bluefly-Hub", repo_name="Cerberus_Sensitivity_Analysis"):
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/releases/latest"
    
    def check_for_updates(self):
        """Check if a newer version is available on GitHub"""
        try:
            with urllib.request.urlopen(self.api_url, timeout=5) as response:
                data = json.loads(response.read().decode())
                latest_version = data['tag_name'].lstrip('v')
                current_version = __version__
                
                if self._is_newer_version(latest_version, current_version):
                    return True, latest_version, data
                return False, current_version, None
        except Exception as e:
            print(f"Could not check for updates: {e}")
            return False, __version__, None
    
    def _is_newer_version(self, latest, current):
        """Compare version strings (e.g., '1.2.0' vs '1.1.0')"""
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            return latest_parts > current_parts
        except:
            return False
    
    def download_and_install_update(self, download_url, latest_version):
        """Download and install the update"""
        try:
            # Only works for compiled .exe files
            if not getattr(sys, 'frozen', False):
                return False
            
            current_exe = sys.executable
            temp_exe = current_exe + '.new'
            backup_exe = current_exe + '.old'
            
            # Download the new version
            print(f"Downloading update v{latest_version}...")
            urllib.request.urlretrieve(download_url, temp_exe)
            
            # Create a batch script to replace the exe after this process exits
            batch_file = os.path.join(os.path.dirname(current_exe), 'update.bat')
            batch_content = f"""@echo off
timeout /t 2 /nobreak > nul
move /y "{current_exe}" "{backup_exe}"
move /y "{temp_exe}" "{current_exe}"
start "" "{current_exe}"
del "%~f0"
"""
            
            with open(batch_file, 'w') as f:
                f.write(batch_content)
            
            # Run the batch file and exit
            subprocess.Popen(['cmd', '/c', batch_file], 
                           creationflags=subprocess.CREATE_NO_WINDOW)
            
            return True
            
        except Exception as e:
            print(f"Error installing update: {e}")
            # Clean up temp files
            if os.path.exists(temp_exe):
                try:
                    os.remove(temp_exe)
                except:
                    pass
            return False


def check_and_update():
    """Check for updates and prompt user to install if available"""
    updater = AutoUpdater()
    has_update, version, release_data = updater.check_for_updates()
    
    if has_update and release_data:
        # Find the .exe asset in the release
        exe_asset = None
        for asset in release_data.get('assets', []):
            if asset['name'].endswith('.exe'):
                exe_asset = asset
                break
        
        if exe_asset:
            root = tk.Tk()
            root.withdraw()
            
            message = f"A new version (v{version}) is available!\n\n"
            message += f"Current version: v{__version__}\n"
            message += f"New version: v{version}\n\n"
            message += "Would you like to update now?"
            
            result = messagebox.askyesno("Update Available", message)
            
            if result:
                download_url = exe_asset['browser_download_url']
                if updater.download_and_install_update(download_url, version):
                    messagebox.showinfo("Update", "Update will be installed when you close this application.")
                    root.destroy()
                    return True
                else:
                    messagebox.showerror("Update Failed", "Could not install update. Please download manually from GitHub.")
            
            root.destroy()
    
    return False


if __name__ == "__main__":
    check_and_update()
