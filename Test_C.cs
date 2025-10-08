using System;
using System.Windows.Automation;

namespace Test_C
{
    class Program
    {
        static void Main(string[] args)
        {
            AutomationElement mainWindow = FindWindowByTitleSubstring("Orpheus");

            if (mainWindow == null)
            {
                Console.WriteLine("Main window not found.");
                return;
            }

            AutomationElement menuBar = mainWindow.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar));

            if (menuBar == null)
            {
                Console.WriteLine("Menu bar not found.");
                return;
            }

            AutomationElement toolsMenu = menuBar.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Tools"));

            if (toolsMenu == null)
            {
                Console.WriteLine("Tools menu not found.");
                return;
            }

            // Open the Tools menu using any supported pattern
            if (toolsMenu.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expandPatternObj))
            {
                ((ExpandCollapsePattern)expandPatternObj).Expand();
            }
            else if (toolsMenu.TryGetCurrentPattern(InvokePattern.Pattern, out object toolsInvokeObj))
            {
                ((InvokePattern)toolsInvokeObj).Invoke();
            }
            else
            {
                Console.WriteLine("Tools menu does not support ExpandCollapse or Invoke.");
                return;
            }

            AutomationElement sensitivityMenu = toolsMenu.FindFirst(
                TreeScope.Descendants,
                new PropertyCondition(AutomationElement.NameProperty, "Sensitivity Analysis..."));

            if (sensitivityMenu == null)
            {
                Console.WriteLine("Sensitivity Analysis menu item not found.");
                return;
            }

            if (sensitivityMenu.TryGetCurrentPattern(InvokePattern.Pattern, out object invokePatternObj))
            {
                ((InvokePattern)invokePatternObj).Invoke();
                Console.WriteLine("Clicked 'Sensitivity Analysis...'");
            }
            else
            {
                Console.WriteLine("Sensitivity Analysis item does not support Invoke.");
            }
        }

        private static AutomationElement FindWindowByTitleSubstring(string titleFragment)
        {
            if (string.IsNullOrWhiteSpace(titleFragment))
            {
                return null;
            }

            Condition windowCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            AutomationElementCollection windows = AutomationElement.RootElement.FindAll(TreeScope.Children, windowCondition);

            foreach (AutomationElement window in windows)
            {
                string windowTitle = window.Current.Name;
                if (!string.IsNullOrEmpty(windowTitle) &&
                    windowTitle.IndexOf(titleFragment, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return window;
                }
            }

            return null;
        }
    }
}