using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Threading;

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

            if (patterns.Any(p => p.Contains("Invoke", StringComparison.OrdinalIgnoreCase)))
            {
                return ButtonAction.Invoke;
            }

            if (patterns.Any(p => p.Contains("SelectionItem", StringComparison.OrdinalIgnoreCase)))
            {
                return ButtonAction.Select;
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
        private string? _lastFieldName;

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
                string continuation = line.Trim();
                if (!string.IsNullOrEmpty(continuation) && !string.IsNullOrEmpty(_lastFieldName))
                {
                    if (_fields.TryGetValue(_lastFieldName, out var previous))
                    {
                        _fields[_lastFieldName] = $"{previous}{Environment.NewLine}{continuation}";
                    }
                }
                return;
            }

            string fieldName = line[..separatorIndex].Trim();
            string fieldValue = line[(separatorIndex + 1)..].Trim();
            if (string.IsNullOrEmpty(fieldName))
            {
                _lastFieldName = null;
                return;
            }

            _fields[fieldName] = fieldValue;
            _lastFieldName = fieldName;
        }

        public string? GetField(string name) =>
            _fields.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value) ? value : null;

        public ButtonDescriptor ToDescriptor()
        {
            string? automationId = AutomationParsers.StripQuotes(GetField("AutomationId") ?? GetField("Automation Id") ?? GetField("AutomationID"));
            string? name = AutomationParsers.StripQuotes(GetField("Name"));
            string? controlTypeRaw = GetField("ControlType") ?? GetField("Control Type");
            ControlType? controlType = AutomationParsers.ParseControlType(controlTypeRaw);
            IReadOnlyList<string> patterns = BuildPatternList();
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

        private IReadOnlyList<string> BuildPatternList()
        {
            List<string> patternList = new();
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

            void AddPattern(string? value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return;
                }

                if (seen.Add(value))
                {
                    patternList.Add(value);
                }
            }

            foreach (string pattern in AutomationParsers.ParsePatterns(GetField("Patterns") ?? GetField("Pattern")))
            {
                AddPattern(pattern);
            }

            foreach (string inferred in AutomationParsers.ParsePatternAvailability(RawLines))
            {
                AddPattern(inferred);
            }

            return patternList;
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

        public static IReadOnlyList<string> ParsePatternAvailability(IEnumerable<string> rawLines)
        {
            List<string> patterns = new();
            HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);

            foreach (string line in rawLines)
            {
                int separatorIndex = line.IndexOf(':');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                string fieldName = line[..separatorIndex].Trim();
                string fieldValue = line[(separatorIndex + 1)..].Trim();

                if (!fieldValue.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!fieldName.StartsWith("Is", StringComparison.OrdinalIgnoreCase) ||
                    !fieldName.EndsWith("PatternAvailable", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string core = fieldName[2..];
                if (core.Length <= "Available".Length)
                {
                    continue;
                }

                core = core[..^"Available".Length];
                if (string.IsNullOrWhiteSpace(core))
                {
                    continue;
                }

                string patternName = core;
                if (!patternName.EndsWith("Pattern", StringComparison.OrdinalIgnoreCase))
                {
                    patternName += "Pattern";
                }

                if (seen.Add(patternName))
                {
                    patterns.Add(patternName);
                }
            }

            return patterns;
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
        private readonly Dictionary<string, IReadOnlyList<string>> _resolvedAncestorCache = new(StringComparer.OrdinalIgnoreCase);
        private string _mainWindowName = string.Empty;
        private string? _currentFileName;

        public AutomationRunner(IReadOnlyDictionary<string, ButtonDescriptor> descriptors, string windowPattern)
        {
            _descriptors = descriptors;
            _windowRegex = new Regex(windowPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
        }

        public void InvokeButton(string key)
        {
            ButtonDescriptor descriptor;
            AutomationElement element = ResolveElement(key, out descriptor);
            ExecuteAction(element, descriptor);
        }

        public void SetValue(string key, string value)
        {
            if (value is null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            ButtonDescriptor descriptor;
            AutomationElement element = ResolveElement(key, out descriptor);

            if (TrySetValue(element, value))
            {
                Console.WriteLine($"Set value '{value}' on '{descriptor.Key}'.");
                return;
            }

            throw new InvalidOperationException($"Unable to set value on '{descriptor.Key}'. Control does not expose a writable ValuePattern.");
        }

        public void PrintPatternDiagnostics(string key)
        {
            ButtonDescriptor descriptor;
            AutomationElement element = ResolveElement(key, out descriptor);

            Console.WriteLine($"Pattern diagnostics for '{descriptor.Key}':");

            try
            {
                AutomationPattern[] supported = element.GetSupportedPatterns();
                if (supported.Length == 0)
                {
                    Console.WriteLine("  No supported patterns reported by provider.");
                }
                else
                {
                    foreach (AutomationPattern pattern in supported.OrderBy(p => p.ProgrammaticName, StringComparer.OrdinalIgnoreCase))
                    {
                        bool available = element.TryGetCurrentPattern(pattern, out _);
                        Console.WriteLine($"  {SimplifyPatternName(pattern.ProgrammaticName)} : {(available ? "available" : "not available")}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Failed to query supported patterns: {ex.Message}");
            }
        }

        private AutomationElement ResolveElement(string key, out ButtonDescriptor descriptor)
        {
            if (!_descriptors.TryGetValue(key, out var foundDescriptor))
            {
                throw new InvalidOperationException($"Button '{key}' not found in repository.");
            }
            descriptor = foundDescriptor;

            if (!descriptor.HasSearchCriteria)
            {
                throw new InvalidOperationException(
                    $"Button '{key}' is missing search metadata (AutomationId/Name/ControlType). Please update the inspect dump.");
            }

            _resolvedAncestorCache.Clear();
            _mainWindowName = string.Empty;
            _currentFileName = null;

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

            return element;
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
                    _mainWindowName = title;
                    _currentFileName = ExtractFileNameFromTitle(title);
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

        private AutomationElement? FindElement(AutomationElement root, ButtonDescriptor descriptor)
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

            AutomationElement searchRoot = GetSearchRoot(root, descriptor) ?? root;

            AutomationElement? element = searchRoot.FindFirst(TreeScope.Descendants, searchCondition);
            if (element is not null)
            {
                return element;
            }

            if (!string.IsNullOrEmpty(descriptor.Name))
            {
                element = FindByNormalizedName(searchRoot, descriptor);
                if (element is not null)
                {
                    return element;
                }
            }

            if (!ReferenceEquals(searchRoot, root))
            {
                element = root.FindFirst(TreeScope.Descendants, searchCondition);
                if (element is not null)
                {
                    return element;
                }
            }

            return AutomationElement.RootElement.FindFirst(TreeScope.Descendants, searchCondition);
        }

        private static AutomationElement? FindByNormalizedName(AutomationElement searchRoot, ButtonDescriptor descriptor)
        {
            Condition baseCondition = descriptor.ControlType is not null
                ? new PropertyCondition(AutomationElement.ControlTypeProperty, descriptor.ControlType)
                : Condition.TrueCondition;

            AutomationElementCollection candidates = searchRoot.FindAll(TreeScope.Descendants, baseCondition);
            foreach (AutomationElement candidate in candidates)
            {
                string actualName = candidate.Current.Name ?? string.Empty;
                if (NameMatches(descriptor.Name!, actualName))
                {
                    return candidate;
                }
            }

            return null;
        }

        private void EnsureAncestorsOpen(AutomationElement root, ButtonDescriptor descriptor)
        {
            IReadOnlyList<string> resolvedAncestors = GetResolvedAncestors(descriptor);
            if (resolvedAncestors.Count == 0)
            {
                return;
            }

            AutomationElement currentRoot = root;
            string mainWindowName = _mainWindowName;
            foreach (string ancestorName in resolvedAncestors)
            {
                if (ShouldSkipAncestor(ancestorName, mainWindowName))
                {
                    continue;
                }

                AutomationElement? ancestor = FindAncestor(currentRoot, ancestorName);
                ancestor ??= FindAncestor(root, ancestorName);
                ancestor ??= FindAncestor(AutomationElement.RootElement, ancestorName);
                if (ancestor is null)
                {
                    continue;
                }

                TryExpandOrInvoke(ancestor);
                currentRoot = ancestor;
            }
        }

        private static AutomationElement? FindAncestor(AutomationElement root, string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern))
            {
                return null;
            }

            bool allowWildcard = ContainsWildcard(pattern);
            try
            {
                if (!allowWildcard)
                {
                    return root.FindFirst(
                        TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.NameProperty, pattern));
                }

                Regex regex = CreateAncestorRegex(pattern);

                string rootName = root.Current.Name ?? string.Empty;
                if (regex.IsMatch(rootName))
                {
                    return root;
                }

                AutomationElementCollection allDescendants = root.FindAll(
                    TreeScope.Descendants,
                    Condition.TrueCondition);

                foreach (AutomationElement element in allDescendants)
                {
                    string candidateName = element.Current.Name ?? string.Empty;
                    if (regex.IsMatch(candidateName))
                    {
                        return element;
                    }
                }
            }
            catch
            {
                // Ignore lookup failures; caller will retry with alternate roots.
            }

            return null;
        }

        private AutomationElement? GetSearchRoot(AutomationElement mainWindow, ButtonDescriptor descriptor)
        {
            IReadOnlyList<string> resolvedAncestors = GetResolvedAncestors(descriptor);
            if (resolvedAncestors.Count == 0)
            {
                return mainWindow;
            }

            string mainWindowName = _mainWindowName;
            AutomationElement current = mainWindow;
            IEnumerable<string> orderedAncestors = resolvedAncestors
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Reverse();

            foreach (string ancestorName in orderedAncestors)
            {
                if (ShouldSkipAncestor(ancestorName, mainWindowName))
                {
                    continue;
                }

                AutomationElement? match = FindAncestor(current, ancestorName)
                    ?? FindAncestor(mainWindow, ancestorName)
                    ?? FindAncestor(AutomationElement.RootElement, ancestorName);

                if (match is not null)
                {
                    current = match;
                }
            }

            return current;
        }

        private static bool ContainsWildcard(string value) =>
            value.IndexOfAny(new[] { '*', '?' }) >= 0;

        private static Regex CreateAncestorRegex(string pattern)
        {
            string escaped = Regex.Escape(pattern);
            escaped = escaped.Replace(@"\.\*", ".*");
            escaped = escaped.Replace(@"\.\?", ".?");
            escaped = escaped.Replace(@"\*", ".*");
            escaped = escaped.Replace(@"\?", ".");
            string anchored = $"^{escaped}$";
            return new Regex(anchored, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private IReadOnlyList<string> GetResolvedAncestors(ButtonDescriptor descriptor)
        {
            if (_resolvedAncestorCache.TryGetValue(descriptor.Key, out var cached))
            {
                return cached;
            }

            List<string> resolved = descriptor.Ancestors
                .Select(ResolveAncestorName)
                .ToList();
            _resolvedAncestorCache[descriptor.Key] = resolved;
            return resolved;
        }

        private string ResolveAncestorName(string ancestorName)
        {
            if (string.IsNullOrWhiteSpace(ancestorName))
            {
                return ancestorName;
            }

            string resolved = ancestorName;

            if (!string.IsNullOrEmpty(_mainWindowName))
            {
                resolved = ReplaceToken(resolved, "{{MAIN_WINDOW}}", _mainWindowName);
                resolved = ReplaceToken(resolved, "{{MAIN_WINDOW_NAME}}", _mainWindowName);
            }

            if (!string.IsNullOrEmpty(_currentFileName))
            {
                resolved = ReplaceToken(resolved, "{{FILE_NAME}}", _currentFileName);
                resolved = ReplaceToken(resolved, "{{FILENAME}}", _currentFileName);
            }

            resolved = resolved.Trim('\r', '\n');

            if (ContainsWildcard(resolved) && !string.IsNullOrEmpty(_mainWindowName))
            {
                Regex regex = CreateAncestorRegex(resolved);
                if (regex.IsMatch(_mainWindowName))
                {
                    return _mainWindowName;
                }
            }

            return resolved;
        }

        private static string ReplaceToken(string input, string token, string replacement)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(token))
            {
                return input;
            }

            return input.Replace(token, replacement, StringComparison.OrdinalIgnoreCase);
        }

        private static string? ExtractFileNameFromTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return null;
            }

            Match match = Regex.Match(title, @"<\s*(.+?)\s*\(local\)\s*>", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            return null;
        }

        private static bool ShouldSkipAncestor(string ancestorName, string mainWindowName)
        {
            if (string.IsNullOrWhiteSpace(ancestorName))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(mainWindowName) &&
                ancestorName.Equals(mainWindowName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (ancestorName.StartsWith("Desktop", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
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

        private static string SimplifyPatternName(string programmaticName)
        {
            if (string.IsNullOrWhiteSpace(programmaticName))
            {
                return "<unknown>";
            }

            string[] parts = programmaticName.Split('.', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return programmaticName;
            }

            string segment = parts[^1];
            if (segment.Equals("Pattern", StringComparison.OrdinalIgnoreCase) && parts.Length >= 2)
            {
                segment = parts[^2];
            }

            segment = segment.Replace("PatternIdentifiers", "Pattern", StringComparison.OrdinalIgnoreCase);
            if (!segment.EndsWith("Pattern", StringComparison.OrdinalIgnoreCase))
            {
                segment += "Pattern";
            }

            return segment;
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

            const int LegacyPatternId = 10018; // UIA_LegacyIAccessiblePatternId
            AutomationPattern legacyPattern = AutomationPattern.LookupById(LegacyPatternId);
            if (legacyPattern is not null && element.TryGetCurrentPattern(legacyPattern, out object legacyObj))
            {
                try
                {
                    MethodInfo? doDefaultAction = legacyObj.GetType().GetMethod("DoDefaultAction", Type.EmptyTypes);
                    if (doDefaultAction is not null)
                    {
                        FocusElementIfPossible(element);
                        doDefaultAction.Invoke(legacyObj, Array.Empty<object>());
                        return true;
                    }

                    MethodInfo? legacyDoDefault = legacyObj.GetType().GetInterface("System.Windows.Automation.Provider.ILegacyIAccessibleProvider")?
                        .GetMethod("DoDefaultAction");
                    if (legacyDoDefault is not null)
                    {
                        FocusElementIfPossible(element);
                        legacyDoDefault.Invoke(legacyObj, Array.Empty<object>());
                        return true;
                    }
                }
                catch (TargetInvocationException)
                {
                    // Provider threw during default action; fall through to failure.
                }
                catch (MethodAccessException)
                {
                    // Provider disallows invocation; fall through.
                }
                catch
                {
                    // Any other reflection errors ignored; fall through.
                }
            }

            if (TryDoubleClick(element))
            {
                return true;
            }

            return false;
        }

        private static void FocusElementIfPossible(AutomationElement element)
        {
            try
            {
                if (element.Current.IsKeyboardFocusable)
                {
                    element.SetFocus();
                }
            }
            catch
            {
                // Focus best-effort; ignore failures.
            }
        }

        private bool TrySetValue(AutomationElement element, string value)
        {
            if (TrySetValueDirect(element, value))
            {
                return true;
            }

            bool needsEdit = false;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out object patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                needsEdit = valuePattern.Current.IsReadOnly;
            }

            if (!needsEdit)
            {
                return false;
            }

            FocusElementIfPossible(element);

            if (!TryDoubleClick(element))
            {
                return false;
            }

            Thread.Sleep(120);

            AutomationElement? focused = AutomationElement.FocusedElement;
            if (focused is not null && !focused.Equals(element) && TrySetValueDirect(focused, value))
            {
                return true;
            }

            AutomationElement? editable = element.FindFirst(
                TreeScope.Subtree,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));

            if (editable is not null && TrySetValueDirect(editable, value))
            {
                return true;
            }

            return TrySetValueDirect(element, value);
        }

        private static bool TrySetValueDirect(AutomationElement element, string value)
        {
            try
            {
                if (!element.TryGetCurrentPattern(ValuePattern.Pattern, out object patternObj))
                {
                    return false;
                }

                var valuePattern = (ValuePattern)patternObj;
                if (valuePattern.Current.IsReadOnly)
                {
                    return false;
                }

                FocusElementIfPossible(element);
                valuePattern.SetValue(value);
                return true;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private static bool TryDoubleClick(AutomationElement element)
        {
            try
            {
                System.Windows.Rect rect = element.Current.BoundingRectangle;
                if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
                {
                    return false;
                }

                int targetX = (int)(rect.X + (rect.Width / 2));
                int targetY = (int)(rect.Y + (rect.Height / 2));

                if (!GetCursorPos(out POINT originalPos))
                {
                    originalPos = new POINT { X = targetX, Y = targetY };
                }

                FocusElementIfPossible(element);

                if (!SetCursorPos(targetX, targetY))
                {
                    return false;
                }

                Thread.Sleep(50);
                SendClick();
                Thread.Sleep(60);
                SendClick();
                Thread.Sleep(80);

                SetCursorPos(originalPos.X, originalPos.Y);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void SendClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

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
                FocusElementIfPossible(element);
                return true;
            }

            return false;
        }

        private static bool NameMatches(string expected, string actual)
        {
            if (string.IsNullOrEmpty(expected))
            {
                return false;
            }

            if (string.Equals(expected, actual, StringComparison.Ordinal))
            {
                return true;
            }

            string normalizedExpected = NormalizeWhitespace(expected);
            string normalizedActual = NormalizeWhitespace(actual);
            return !string.IsNullOrEmpty(normalizedExpected) &&
                   string.Equals(normalizedExpected, normalizedActual, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeWhitespace(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return Regex.Replace(value, @"\s+", " ").Trim();
        }
    }

    internal static class Program
    {
        private const string DefaultWindowRegex = ".*(Orpheus|Cerberus).*";
        private const string DefaultDumpRelativePath = "inspect_dumps\\Windows_Inspect_Dump.txt";

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
                case "patterns":
                case "diagnose":
                case "set":
                    if (arguments.Count == 0)
                    {
                        throw new InvalidOperationException("Missing button key for invoke command.");
                    }

                    string key = arguments.Dequeue();
                    var runner = new AutomationRunner(descriptors, windowRegex);
                    if (command is "patterns" or "diagnose")
                    {
                        runner.PrintPatternDiagnostics(key);
                    }
                    else if (command is "set")
                    {
                        if (arguments.Count == 0)
                        {
                            throw new InvalidOperationException("Missing value for set command.");
                        }

                        string newValue = arguments.Dequeue();
                        runner.SetValue(key, newValue);
                    }
                    else
                    {
                        runner.InvokeButton(key);
                    }
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
            Console.WriteLine("  Drill_Down_With_C.exe [--dump <path>] [--window-regex <regex>] list");
            Console.WriteLine("  Drill_Down_With_C.exe [--dump <path>] [--window-regex <regex>] invoke <button-key>");
            Console.WriteLine("  Drill_Down_With_C.exe [--dump <path>] [--window-regex <regex>] patterns <button-key>");
            Console.WriteLine("  Drill_Down_With_C.exe [--dump <path>] [--window-regex <regex>] set <button-key> <value>");
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
