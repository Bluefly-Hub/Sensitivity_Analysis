using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Automation;

#nullable enable

namespace Cerberus.ButtonAutomation
{
    internal enum ButtonAction
    {
        Default,
        Invoke,
        Toggle,
        Select,
    }

    internal sealed class ButtonDescriptor
    {
        public ButtonDescriptor(
            string key,
            string? automationId,
            string? name,
            ControlType? controlType,
            IReadOnlyList<string> patterns,
            bool? isEnabled,
            IReadOnlyList<string> rawDump,
            IReadOnlyList<string> ancestors)
        {
            Key = key;
            AutomationId = string.IsNullOrWhiteSpace(automationId) ? null : automationId;
            Name = string.IsNullOrWhiteSpace(name) ? null : name;
            ControlType = controlType;
            Patterns = patterns;
            IsEnabled = isEnabled;
            RawDump = rawDump;
            Ancestors = ancestors;
            PreferredAction = DeterminePreferredAction(controlType, patterns);
        }

        public string Key { get; }
        public string? AutomationId { get; }
        public string? Name { get; }
        public ControlType? ControlType { get; }
        public IReadOnlyList<string> Patterns { get; }
        public bool? IsEnabled { get; }
        public IReadOnlyList<string> RawDump { get; }
        public IReadOnlyList<string> Ancestors { get; }
        public ButtonAction PreferredAction { get; }

        public bool HasSearchCriteria =>
            !string.IsNullOrWhiteSpace(AutomationId) ||
            !string.IsNullOrWhiteSpace(Name) ||
            ControlType is not null;

        private static ButtonAction DeterminePreferredAction(ControlType? controlType, IReadOnlyList<string> patterns)
        {
            if (controlType == ControlType.CheckBox || patterns.Any(p => p.Contains("Toggle", StringComparison.OrdinalIgnoreCase)))
            {
                return ButtonAction.Toggle;
            }

            if (patterns.Any(p => p.Contains("SelectionItem", StringComparison.OrdinalIgnoreCase)))
            {
                return ButtonAction.Select;
            }

            if (patterns.Any(p => p.Contains("Invoke", StringComparison.OrdinalIgnoreCase)))
            {
                return ButtonAction.Invoke;
            }

            if (controlType == ControlType.Button || controlType == ControlType.MenuItem)
            {
                return ButtonAction.Invoke;
            }

            return ButtonAction.Default;
        }
    }

    internal sealed class InspectDumpEntry
    {
        private readonly Dictionary<string, string> _fields = new(StringComparer.OrdinalIgnoreCase);

        public InspectDumpEntry(string key)
        {
            Key = key;
        }

        public string Key { get; }
        public IList<string> RawLines { get; } = new List<string>();

        public void AddLine(string line)
        {
            RawLines.Add(line);
            int separatorIndex = line.IndexOf(':');
            if (separatorIndex <= 0)
            {
                return;
            }

            string fieldName = line[..separatorIndex].Trim();
            string fieldValue = line[(separatorIndex + 1)..].Trim();
            if (string.IsNullOrEmpty(fieldName))
            {
                return;
            }

            _fields[fieldName] = fieldValue;
        }

        public string? GetField(string name) =>
            _fields.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value) ? value : null;

        public ButtonDescriptor ToDescriptor()
        {
            string? automationId = AutomationParsers.StripQuotes(GetField("AutomationId") ?? GetField("Automation Id") ?? GetField("AutomationID"));
            string? name = AutomationParsers.StripQuotes(GetField("Name"));
            string? controlTypeRaw = GetField("ControlType") ?? GetField("Control Type");
            ControlType? controlType = AutomationParsers.ParseControlType(controlTypeRaw);
            IReadOnlyList<string> patterns = AutomationParsers.ParsePatterns(GetField("Patterns") ?? GetField("Pattern"));
            bool? isEnabled = AutomationParsers.ParseNullableBool(GetField("IsEnabled") ?? GetField("Is Enabled"));
            IReadOnlyList<string> ancestors = AutomationParsers.ParseAncestors(RawLines);

            return new ButtonDescriptor(
                Key,
                automationId,
                name,
                controlType,
                patterns,
                isEnabled,
                RawLines.ToList(),
                ancestors);
        }
    }

    internal static class AutomationParsers
    {
        private static readonly Dictionary<string, ControlType> ControlTypeMap = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Button"] = ControlType.Button,
            ["ControlType.Button"] = ControlType.Button,
            ["MenuItem"] = ControlType.MenuItem,
            ["ControlType.MenuItem"] = ControlType.MenuItem,
            ["CheckBox"] = ControlType.CheckBox,
            ["ControlType.CheckBox"] = ControlType.CheckBox,
            ["RadioButton"] = ControlType.RadioButton,
            ["ControlType.RadioButton"] = ControlType.RadioButton,
            ["Hyperlink"] = ControlType.Hyperlink,
            ["ControlType.Hyperlink"] = ControlType.Hyperlink,
        };

        public static ControlType? ParseControlType(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string trimmed = value.Trim();
            if (ControlTypeMap.TryGetValue(trimmed, out var mapped))
            {
                return mapped;
            }

            string simplified = trimmed.Replace("ControlType.", string.Empty, StringComparison.OrdinalIgnoreCase);
            if (ControlTypeMap.TryGetValue(simplified, out var simplifiedMapped))
            {
                return simplifiedMapped;
            }

            foreach ((string key, ControlType controlType) in ControlTypeMap)
            {
                if (trimmed.Contains(key, StringComparison.OrdinalIgnoreCase))
                {
                    return controlType;
                }
            }

            return null;
        }

        public static IReadOnlyList<string> ParsePatterns(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Array.Empty<string>();
            }

            string[] separators = { ",", ";", "|", Environment.NewLine };
            string[] tokens = value.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            return tokens
                .Select(token => token.Trim())
                .Where(token => !string.IsNullOrEmpty(token))
                .ToArray();
        }

        public static bool? ParseNullableBool(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string trimmed = value.Trim();
            if (bool.TryParse(trimmed, out bool result))
            {
                return result;
            }

            if (string.Equals(trimmed, "1", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "yes", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "on", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(trimmed, "0", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "no", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(trimmed, "off", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return null;
        }

        public static string? StripQuotes(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            string trimmed = value.Trim();
            if (trimmed.Length >= 2 && trimmed.StartsWith("\"", StringComparison.Ordinal) && trimmed.EndsWith("\"", StringComparison.Ordinal))
            {
                return trimmed[1..^1];
            }

            return trimmed;
        }

        public static IReadOnlyList<string> ParseAncestors(IEnumerable<string> rawLines)
        {
            var ancestors = new List<string>();
            bool inAncestors = false;
            foreach (string line in rawLines)
            {
                if (line.StartsWith("Ancestors:", StringComparison.OrdinalIgnoreCase))
                {
                    inAncestors = true;
                    string remainder = line[10..].Trim();
                    string? name = StripQuotes(ExtractAncestorName(remainder));
                    if (!string.IsNullOrEmpty(name))
                    {
                        ancestors.Add(name);
                    }
                    continue;
                }

                if (inAncestors)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        break;
                    }

                    string trimmed = line.Trim();
                    string? name = StripQuotes(ExtractAncestorName(trimmed));
                    if (!string.IsNullOrEmpty(name))
                    {
                        ancestors.Add(name);
                    }
                }
            }

            return ancestors;
        }

        private static string? ExtractAncestorName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            if (value.StartsWith("[", StringComparison.Ordinal))
            {
                return null;
            }

            int quoteStart = value.IndexOf('"');
            if (quoteStart >= 0)
            {
                int quoteEnd = value.IndexOf('"', quoteStart + 1);
                if (quoteEnd > quoteStart)
                {
                    return value.Substring(quoteStart, quoteEnd - quoteStart + 1);
                }
            }

            int spaceIndex = value.IndexOf(' ');
            return spaceIndex > 0 ? value[..spaceIndex] : value;
        }
    }

    internal static class InspectDumpRepository
    {
        public static IReadOnlyDictionary<string, ButtonDescriptor> Load(string path)
        {
            var descriptors = new Dictionary<string, ButtonDescriptor>(StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(path))
            {
                Console.Error.WriteLine($"Inspect dump file not found at '{path}'.");
                return descriptors;
            }

            InspectDumpEntry? currentEntry = null;

            foreach (string line in File.ReadLines(path))
            {
                string trimmed = line.Trim();

                if (trimmed.StartsWith("#", StringComparison.Ordinal) || string.IsNullOrEmpty(trimmed))
                {
                    continue;
                }

                if (trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal))
                {
                    if (currentEntry is not null)
                    {
                        ButtonDescriptor descriptor = currentEntry.ToDescriptor();
                        descriptors[descriptor.Key] = descriptor;
                    }

                    string key = trimmed[1..^1].Trim();
                    if (!string.IsNullOrEmpty(key) && IsValidKey(key))
                    {
                        currentEntry = new InspectDumpEntry(key);
                    }
                    else
                    {
                        currentEntry = null;
                    }

                    continue;
                }

                currentEntry?.AddLine(line);
            }

            if (currentEntry is not null)
            {
                ButtonDescriptor descriptor = currentEntry.ToDescriptor();
                descriptors[descriptor.Key] = descriptor;
            }

            return descriptors;
        }

        private static bool IsValidKey(string key) =>
            key.Length > 0 && key.All(c => char.IsLetterOrDigit(c) || c is '_' or '-');
    }

    internal sealed class AutomationRunner
    {
        private readonly IReadOnlyDictionary<string, ButtonDescriptor> _descriptors;
        private readonly Regex _windowRegex;

        public AutomationRunner(IReadOnlyDictionary<string, ButtonDescriptor> descriptors, string windowPattern)
        {
            _descriptors = descriptors;
            _windowRegex = new Regex(windowPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        }

        public void InvokeButton(string key)
        {
            if (!_descriptors.TryGetValue(key, out var descriptor))
            {
                throw new InvalidOperationException($"Button '{key}' not found in repository.");
            }

            if (!descriptor.HasSearchCriteria)
            {
                throw new InvalidOperationException(
                    $"Button '{key}' is missing search metadata (AutomationId/Name/ControlType). Please update the inspect dump.");
            }

            AutomationElement mainWindow = FindMainWindow();
            FocusWindow(mainWindow);
            AutomationElement? element = FindElement(mainWindow, descriptor);

            if (element is null)
            {
                EnsureAncestorsOpen(mainWindow, descriptor);
                element = FindElement(mainWindow, descriptor);
            }

            if (element is null)
            {
                throw new InvalidOperationException(
                    $"Unable to locate UI Automation element for button '{key}'. Verify the inspect dump metadata.");
            }

            if (!element.Current.IsEnabled)
            {
                throw new InvalidOperationException($"Button '{key}' is currently disabled in the UI.");
            }

            if (descriptor.IsEnabled.HasValue && descriptor.IsEnabled.Value == false)
            {
                Console.WriteLine($"Warning: Inspect dump indicates '{key}' was disabled. Proceeding anyway.");
            }

            ExecuteAction(element, descriptor);
        }

        private AutomationElement FindMainWindow()
        {
            Condition windowCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            AutomationElementCollection windows = AutomationElement.RootElement.FindAll(TreeScope.Children, windowCondition);

            foreach (AutomationElement window in windows)
            {
                string title = window.Current.Name ?? string.Empty;
                if (_windowRegex.IsMatch(title))
                {
                    return window;
                }
            }

            throw new InvalidOperationException($"Main window matching pattern '{_windowRegex}' was not found.");
        }

        private static void FocusWindow(AutomationElement window)
        {
            try
            {
                if (window.Current.IsKeyboardFocusable)
                {
                    window.SetFocus();
                }

                if (window.TryGetCurrentPattern(WindowPattern.Pattern, out object patternObj))
                {
                    var pattern = (WindowPattern)patternObj;
                    if (pattern.Current.WindowVisualState == WindowVisualState.Minimized)
                    {
                        pattern.SetWindowVisualState(WindowVisualState.Normal);
                    }
                }
            }
            catch
            {
                // Non-fatal; best-effort focus.
            }
        }

        private static AutomationElement? FindElement(AutomationElement root, ButtonDescriptor descriptor)
        {
            var conditions = new List<Condition>();

            if (!string.IsNullOrEmpty(descriptor.AutomationId))
            {
                conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, descriptor.AutomationId));
            }

            if (!string.IsNullOrEmpty(descriptor.Name))
            {
                conditions.Add(new PropertyCondition(AutomationElement.NameProperty, descriptor.Name));
            }

            if (descriptor.ControlType is not null)
            {
                conditions.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, descriptor.ControlType));
            }

            Condition searchCondition = conditions.Count switch
            {
                0 => Condition.TrueCondition,
                1 => conditions[0],
                _ => new AndCondition(conditions.ToArray()),
            };

            AutomationElement? element = root.FindFirst(TreeScope.Descendants, searchCondition);
            return element ?? AutomationElement.RootElement.FindFirst(TreeScope.Descendants, searchCondition);
        }

        private static void EnsureAncestorsOpen(AutomationElement root, ButtonDescriptor descriptor)
        {
            if (descriptor.Ancestors.Count == 0)
            {
                return;
            }

            AutomationElement currentRoot = root;
            foreach (string ancestorName in descriptor.Ancestors)
            {
                if (string.IsNullOrWhiteSpace(ancestorName))
                {
                    continue;
                }

                AutomationElement? ancestor = FindAncestor(currentRoot, ancestorName);
                ancestor ??= FindAncestor(AutomationElement.RootElement, ancestorName);
                if (ancestor is null)
                {
                    continue;
                }

                TryExpandOrInvoke(ancestor);
                currentRoot = ancestor;
            }
        }

        private static AutomationElement? FindAncestor(AutomationElement root, string name)
        {
            try
            {
                return root.FindFirst(
                    TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, name));
            }
            catch
            {
                return null;
            }
        }

        private static void TryExpandOrInvoke(AutomationElement element)
        {
            try
            {
                if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expandObj))
                {
                    var pattern = (ExpandCollapsePattern)expandObj;
                    if (pattern.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
                    {
                        pattern.Expand();
                        System.Threading.Thread.Sleep(150);
                    }
                    return;
                }

                if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object invokeObj))
                {
                    ((InvokePattern)invokeObj).Invoke();
                    System.Threading.Thread.Sleep(150);
                }
            }
            catch
            {
                // Swallow and continue; ancestor invocation is best-effort.
            }
        }

        private static void ExecuteAction(AutomationElement element, ButtonDescriptor descriptor)
        {
            foreach (ButtonAction action in BuildActionOrder(descriptor.PreferredAction))
            {
                if (TryExecute(element, action))
                {
                    Console.WriteLine($"Executed {action} on '{descriptor.Key}'.");
                    return;
                }
            }

            string availablePatterns = string.Join(", ", descriptor.Patterns);
            throw new InvalidOperationException(
                $"No supported automation pattern found for '{descriptor.Key}'. Patterns from dump: {availablePatterns}");
        }

        private static IEnumerable<ButtonAction> BuildActionOrder(ButtonAction preferred)
        {
            return preferred switch
            {
                ButtonAction.Invoke => new[] { ButtonAction.Invoke, ButtonAction.Select, ButtonAction.Toggle, ButtonAction.Default },
                ButtonAction.Toggle => new[] { ButtonAction.Toggle, ButtonAction.Invoke, ButtonAction.Select, ButtonAction.Default },
                ButtonAction.Select => new[] { ButtonAction.Select, ButtonAction.Invoke, ButtonAction.Toggle, ButtonAction.Default },
                _ => new[] { ButtonAction.Invoke, ButtonAction.Select, ButtonAction.Toggle, ButtonAction.Default },
            };
        }

        private static bool TryExecute(AutomationElement element, ButtonAction action)
        {
            try
            {
                return action switch
                {
                    ButtonAction.Invoke => TryInvoke(element),
                    ButtonAction.Toggle => TryToggle(element),
                    ButtonAction.Select => TrySelect(element),
                    _ => false,
                };
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private static bool TryInvoke(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object pattern))
            {
                ((InvokePattern)pattern).Invoke();
                return true;
            }

            return false;
        }

        private static bool TryToggle(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(TogglePattern.Pattern, out object pattern))
            {
                ((TogglePattern)pattern).Toggle();
                return true;
            }

            return false;
        }

        private static bool TrySelect(AutomationElement element)
        {
            if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern))
            {
                ((SelectionItemPattern)pattern).Select();
                return true;
            }

            return false;
        }
    }

    internal static class Program
    {
        private const string DefaultWindowRegex = ".*(Orpheus|Cerberus).*";
        private const string DefaultDumpRelativePath = "inspect_dumps\\button_dump_template.txt";

        private static readonly string ProjectRoot = ResolveProjectRoot();

        private static int Main(string[] args)
        {
            try
            {
                int exitCode = Run(args);
                Environment.Exit(exitCode);
                return exitCode;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Environment.Exit(1);
                return 1;
            }
        }

        private static int Run(string[] args)
        {
            if (args.Length == 0)
            {
                PrintUsage();
                return 1;
            }

            string dumpPath = GetDefaultDumpPath();
            string windowRegex = DefaultWindowRegex;

            var arguments = new Queue<string>(args);
            while (arguments.Count > 0 && arguments.Peek().StartsWith("--", StringComparison.Ordinal))
            {
                string option = arguments.Dequeue();
                switch (option)
                {
                    case "--dump":
                        if (arguments.Count == 0)
                        {
                            throw new InvalidOperationException("Missing value for --dump option.");
                        }

                        dumpPath = ResolvePath(arguments.Dequeue());
                        break;

                    case "--window-regex":
                        if (arguments.Count == 0)
                        {
                            throw new InvalidOperationException("Missing value for --window-regex option.");
                        }

                        windowRegex = arguments.Dequeue();
                        break;

                    default:
                        throw new InvalidOperationException($"Unknown option '{option}'.");
                }
            }

            if (arguments.Count == 0)
            {
                PrintUsage();
                return 1;
            }

            string command = arguments.Dequeue().ToLowerInvariant();
            IReadOnlyDictionary<string, ButtonDescriptor> descriptors = InspectDumpRepository.Load(dumpPath);

            if (descriptors.Count == 0)
            {
                Console.Error.WriteLine("No button entries were loaded from the inspect dump.");
                return 1;
            }

            switch (command)
            {
                case "list":
                    PrintRepository(descriptors);
                    return 0;

                case "invoke":
                case "press":
                case "run":
                    if (arguments.Count == 0)
                    {
                        throw new InvalidOperationException("Missing button key for invoke command.");
                    }

                    string key = arguments.Dequeue();
                    var runner = new AutomationRunner(descriptors, windowRegex);
                    runner.InvokeButton(key);
                    return 0;

                default:
                    throw new InvalidOperationException($"Unknown command '{command}'.");
            }
        }

        private static void PrintRepository(IReadOnlyDictionary<string, ButtonDescriptor> descriptors)
        {
            Console.WriteLine("Available buttons:");
            foreach (ButtonDescriptor descriptor in descriptors.Values.OrderBy(d => d.Key, StringComparer.OrdinalIgnoreCase))
            {
                string idPart = string.IsNullOrEmpty(descriptor.AutomationId) ? "AutomationId=<missing>" : $"AutomationId={descriptor.AutomationId}";
                string namePart = string.IsNullOrEmpty(descriptor.Name) ? "Name=<missing>" : $"Name={descriptor.Name}";
                string typePart = descriptor.ControlType is null ? "ControlType=<unknown>" : $"ControlType={descriptor.ControlType.ProgrammaticName}";
                Console.WriteLine($"  - {descriptor.Key}: {idPart}, {namePart}, {typePart}");
            }
        }

        private static void PrintUsage()
        {
            string defaultDump = GetDefaultDumpPath();
            Console.WriteLine("Button automation helper");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  Test_C.exe [--dump <path>] [--window-regex <regex>] list");
            Console.WriteLine("  Test_C.exe [--dump <path>] [--window-regex <regex>] invoke <button-key>");
            Console.WriteLine();
            Console.WriteLine($"Default inspect dump path: {defaultDump}");
            Console.WriteLine($"Default window regex: {DefaultWindowRegex}");
        }

        private static string GetDefaultDumpPath()
        {
            return Path.Combine(ProjectRoot, DefaultDumpRelativePath);
        }

        private static string ResolvePath(string path)
        {
            if (Path.IsPathRooted(path))
            {
                return path;
            }

            return Path.GetFullPath(Path.Combine(ProjectRoot, path));
        }

        private static string ResolveProjectRoot()
        {
            string baseDir = AppContext.BaseDirectory;
            string projectRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", ".."));
            return projectRoot;
        }
    }
}
