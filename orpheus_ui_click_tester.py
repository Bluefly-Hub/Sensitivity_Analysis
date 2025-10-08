import uiautomation as uia
import subprocess, sys
import time
import pythoncom
from pywinauto import Application, Desktop
from pywinauto.uia_defines import NoPatternInterfaceError

# Make misses fast
uia.SetGlobalSearchTimeout(0.3)


def Sensitivity_Analysis():
    from pywinauto.application import Application

    app = Application(backend='uia').connect(title_re='Orpheus*')
    dlg = app.window(title_re='Orpheus*')

    # If it’s a true Win32/WinForms menu, this often just works:
    # Adjust menu path to the real one that contains the item
    dlg.menu_select("Tools->Sensitivity Analysis*")
    print("Sensitivity Analysis opened")

def Sensitivity_Analysis2():
    """ Click Sensitivity Analysis menu item in Orpheus main window """
    # 1) Attach to your Orpheus window (adjust RegexName if needed)
    Main_Window = uia.WindowControl(RegexName=r'Orpheus*')

    # Many WPF/WinForms apps keep menu items in the tree even when the submenu isn't dropped.
    target = uia.MenuItemControl(searchFromControl=Main_Window, Name='Sensitivity Analysis...')
    inv = target.GetInvokePattern()
    #try:
    inv.Invoke()
    inv = None





def exit():
    # Attach to the main window
    Main_Window = uia.WindowControl(RegexName=r'Orpheus*', Depth=1)

    # Drill directly to the Exit button by AutomationId
    exit_btn = Main_Window.ButtonControl(AutomationId="cmdExit")

    # Invoke the button
    exit_btn.GetInvokePattern().Invoke()

if __name__ == "__main__":
    Sensitivity_Analysis()
    exit()