"""
Main entry point for the Cerberus Sensitivity Analysis Application
Run this file to start the GUI
"""
import tkinter as tk
from GUI_Automation import launch_gui
from updater import check_and_update
from version import __version__


def main():
    """Start the automation GUI application"""
    # Check for updates first
    print(f"Cerberus Sensitivity Analysis v{__version__}")
    if check_and_update():
        # Update is being installed, exit the application
        return
    
    # Start the GUI
    launch_gui()


if __name__ == "__main__":
    main()
