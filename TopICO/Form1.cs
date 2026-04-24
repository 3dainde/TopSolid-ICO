using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace TopICO;

public partial class Form1 : Form
{
    private const string LegacyTopAutoProgId = "top_auto_file";
    private const string LegacyDftAutoProgId = "dft_auto_file";
    private const string PreferredTopProgId = "TopSolid.SolidDocument";
    private const string PreferredDftProgId = "TopSolid.DraftDocument";

    private Label lblTitle = null!;
    private Label lblStatus = null!;
    private Label lblVersion = null!;
    private Button btnAnalyser = null!;
    private Button btnModifier = null!;
    private RichTextBox txtLog = null!;
    private Panel panelBorder = null!;

    public Form1()
    {
        InitializeComponent();
        BuildUI();
        RunAnalysis();
    }

    private void BuildUI()
    {
        Text = "Top'ICO";
        ClientSize = new Size(420, 550);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.White;

        // Red border panel
        panelBorder = new Panel
        {
            Location = new Point(6, 6),
            Size = new Size(ClientSize.Width - 12, ClientSize.Height - 12),
            BackColor = Color.FromArgb(220, 50, 20),
        };
        Controls.Add(panelBorder);

        var panelInner = new Panel
        {
            Location = new Point(3, 3),
            Size = new Size(panelBorder.Width - 6, panelBorder.Height - 6),
            BackColor = Color.White,
        };
        panelBorder.Controls.Add(panelInner);

        // TopSolid logo text
        var lblLogo = new Label
        {
            Text = "TopSolid",
            Font = new Font("Segoe UI", 14, FontStyle.Italic | FontStyle.Bold),
            ForeColor = Color.FromArgb(220, 50, 20),
            AutoSize = true,
            Location = new Point(8, 6),
        };
        panelInner.Controls.Add(lblLogo);

        // Status label
        lblStatus = new Label
        {
            Text = "Analyse d'installation",
            Font = new Font("Segoe UI", 11, FontStyle.Regular),
            ForeColor = Color.Black,
            Location = new Point(8, 40),
            Size = new Size(panelInner.Width - 16, 28),
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        panelInner.Controls.Add(lblStatus);

        // Log area
        txtLog = new RichTextBox
        {
            Location = new Point(8, 75),
            Size = new Size(panelInner.Width - 16, 270),
            ReadOnly = true,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.FromArgb(250, 250, 250),
            Font = new Font("Consolas", 8.5f),
            ScrollBars = RichTextBoxScrollBars.Vertical,
        };
        panelInner.Controls.Add(txtLog);

        // Analyser button
        int btnWidth = 170;
        int btnSpacing = 10;
        int totalWidth = btnWidth * 2 + btnSpacing;
        int startX = (panelInner.Width - totalWidth) / 2;

        btnAnalyser = new Button
        {
            Text = "Analyser",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            Size = new Size(btnWidth, 50),
            Location = new Point(startX, 355),
            FlatStyle = FlatStyle.Standard,
            BackColor = Color.FromArgb(240, 240, 240),
        };
        btnAnalyser.Click += BtnAnalyser_Click;
        panelInner.Controls.Add(btnAnalyser);

        // Modifier button
        btnModifier = new Button
        {
            Text = "Modifier",
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            Size = new Size(btnWidth, 50),
            Location = new Point(startX + btnWidth + btnSpacing, 355),
            FlatStyle = FlatStyle.Standard,
            BackColor = Color.FromArgb(240, 240, 240),
            Enabled = false,
        };
        btnModifier.Click += BtnModifier_Click;
        panelInner.Controls.Add(btnModifier);

        // Version label
        lblVersion = new Label
        {
            Text = $"v{Application.ProductVersion}",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(220, 50, 20),
            AutoSize = true,
        };
        lblVersion.Location = new Point(panelInner.Width - lblVersion.PreferredWidth - 10, panelInner.Height - 28);
        panelInner.Controls.Add(lblVersion);
    }

    private readonly record struct V6Installation(string FolderName, string FullPath, int VersionNumber);

    private List<V6Installation> _detectedVersions = new();
    private string? _bestInstallPath;

    private void BtnAnalyser_Click(object? sender, EventArgs e)
    {
        RunAnalysis();
    }

    private void RunAnalysis()
    {
        txtLog.Clear();
        _detectedVersions.Clear();
        _bestInstallPath = null;
        btnModifier.Enabled = false;

        string misslerRoot = @"C:\Missler\";

        if (!Directory.Exists(misslerRoot))
        {
            Log("Dossier C:\\Missler\\ introuvable.", Color.Red);
            lblStatus.Text = "Aucune installation trouvée";
            return;
        }

        Log("Scan de C:\\Missler\\...", Color.Black);

        var dirs = Directory.GetDirectories(misslerRoot);
        var versionPattern = new Regex(@"^V(6\d{2})$", RegexOptions.IgnoreCase);

        foreach (var dir in dirs.OrderByDescending(d => d))
        {
            string folderName = Path.GetFileName(dir);
            var match = versionPattern.Match(folderName);
            if (match.Success)
            {
                int verNum = int.Parse(match.Groups[1].Value);
                if (verNum < 610 || verNum > 629) continue;
                bool hasExe = FindTopSolidExecutable(dir, verNum) != null;

                var install = new V6Installation(folderName, dir, verNum);
                _detectedVersions.Add(install);

                string status = hasExe ? "OK" : "EXE manquant";
                Color color = hasExe ? Color.DarkGreen : Color.DarkOrange;
                Log($"  {folderName} - {status}", color);
            }
        }

        if (_detectedVersions.Count == 0)
        {
            Log("Aucune version V6xx trouvée.", Color.Red);
            lblStatus.Text = "Aucune version V6 détectée";
            return;
        }

        // Find the best (latest) version that actually exists
        var best = _detectedVersions.OrderByDescending(v => v.VersionNumber).First();
        _bestInstallPath = best.FullPath;

        Log("", Color.Black);
        Log($"Version la plus récente : {best.FolderName}", Color.DarkBlue);
        Log($"Chemin : {best.FullPath}", Color.DarkBlue);
        string? detectedExe = GetTopSolidExePath(best);
        Log($"Exécutable détecté : {detectedExe ?? "introuvable"}", detectedExe == null ? Color.Red : Color.DarkBlue);

        // Analyze registry
        Log("", Color.Black);
        Log("Analyse du registre (chemins d'icônes)...", Color.Black);
        Log("HKLM\\SOFTWARE\\Missler Software", Color.DarkBlue);
        Log("HKLM\\SOFTWARE\\WOW6432Node\\Missler Software", Color.DarkBlue);
        Log("HKCR (associations fichiers TopSolid)", Color.DarkBlue);
        Log("", Color.Black);
        AnalyzeRegistry(best);

        lblStatus.Text = $"Détecté : {best.FolderName}";
        btnModifier.Enabled = true;
    }

    private void AnalyzeRegistry(V6Installation best)
    {
        int issues = 0;
        int totalFound = 0;

        // Check HKLM\SOFTWARE\Missler Software for TopSolid icon paths
        string[] registryRoots = new[]
        {
            @"SOFTWARE\Missler Software",
            @"SOFTWARE\WOW6432Node\Missler Software",
        };

        foreach (string root in registryRoots)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(root);
                if (key == null)
                {
                    Log($"  (clé {root} absente)", Color.Gray);
                    continue;
                }

                int before = issues;
                issues += ScanKeyForIconIssues(key, root, best, ref totalFound);
            }
            catch (Exception ex)
            {
                Log($"  Erreur lecture {root}: {ex.Message}", Color.Red);
            }
        }

        // Check HKCR for TopSolid file associations
        Log("", Color.Black);
        Log("Associations de fichiers (HKCR)...", Color.Black);
        try
        {
            issues += ScanClassesRootForIconIssues(best, ref totalFound);
            issues += AnalyzeFileAssociationIssues(best);
        }
        catch (Exception ex)
        {
            Log($"  Erreur HKCR: {ex.Message}", Color.Red);
        }

        Log("", Color.Black);
        Log($"Résumé : {totalFound} chemin(s) trouvé(s) dans le registre.", Color.DarkBlue);
        if (issues > 0)
        {
            Log($"{issues} chemin(s) d'icône(s) à corriger.", Color.FromArgb(200, 100, 0));
        }
        else
        {
            Log("Tous les chemins d'icônes sont corrects.", Color.DarkGreen);
        }
    }

    private int ScanKeyForIconIssues(RegistryKey parentKey, string parentPath, V6Installation best, ref int totalFound)
    {
        int issues = 0;

        foreach (string subName in parentKey.GetSubKeyNames())
        {
            try
            {
                using var sub = parentKey.OpenSubKey(subName);
                if (sub == null) continue;

                issues += CheckKeyValues(sub, $"{parentPath}\\{subName}", best, ref totalFound);
                issues += ScanKeyForIconIssues(sub, $"{parentPath}\\{subName}", best, ref totalFound);
            }
            catch { }
        }

        return issues;
    }

    private int CheckKeyValues(RegistryKey key, string path, V6Installation best, ref int totalFound)
    {
        int issues = 0;
        var vPattern = new Regex(@"C:\\Missler\\V\d{3}\\", RegexOptions.IgnoreCase);

        foreach (string valueName in key.GetValueNames())
        {
            var val = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (val is not string strVal) continue;

            if (!vPattern.IsMatch(strVal)) continue;

            totalFound++;
            string referenced = vPattern.Match(strVal).Value.TrimEnd('\\');
            string refFolder = Path.GetFileName(referenced);
            string displayName = string.IsNullOrEmpty(valueName) ? "(Default)" : valueName;

            if (strVal.Contains(best.FolderName, StringComparison.OrdinalIgnoreCase))
            {
                Log($"  OK : {path}\\{displayName}", Color.DarkGreen);
                Log($"       -> {strVal}", Color.Gray);
            }
            else if (!Directory.Exists(referenced))
            {
                Log($"  INVALIDE : {path}\\{displayName}", Color.Red);
                Log($"       -> {strVal}", Color.Red);
                Log($"       ({refFolder} introuvable, sera corrigé vers {best.FolderName})", Color.FromArgb(200, 100, 0));
                issues++;
            }
            else
            {
                Log($"  ANCIEN : {path}\\{displayName}", Color.DarkOrange);
                Log($"       -> {strVal}", Color.DarkOrange);
                Log($"       ({refFolder} existe mais n'est pas la version la plus récente)", Color.Gray);
            }
        }

        return issues;
    }

    private int ScanClassesRootForIconIssues(V6Installation best, ref int totalFound)
    {
        int issues = 0;
        var vPattern = new Regex(@"C:\\Missler\\V\d{3}\\", RegexOptions.IgnoreCase);

        // Scan HKCR for TopSolid/Missler related entries
        string[] prefixes = { "TopSolid", "Missler", ".top", ".tpd", ".tgp", ".tpr", ".tpp", ".dft", ".igs", ".stp", ".cam", ".mld", ".wod" };

        using var hkcr = Registry.ClassesRoot;
        foreach (string subName in hkcr.GetSubKeyNames())
        {
            bool relevant = false;
            foreach (var prefix in prefixes)
            {
                if (subName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                    subName.Equals(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    relevant = true;
                    break;
                }
            }
            if (!relevant) continue;

            try
            {
                using var sub = hkcr.OpenSubKey(subName);
                if (sub == null) continue;

                issues += ScanHKCRSubKey(sub, subName, best, vPattern, ref totalFound);
            }
            catch { }
        }

        return issues;
    }

    private int ScanHKCRSubKey(RegistryKey key, string path, V6Installation best, Regex vPattern, ref int totalFound)
    {
        int issues = 0;

        foreach (string valueName in key.GetValueNames())
        {
            var val = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (val is not string strVal) continue;

            if (!vPattern.IsMatch(strVal)) continue;

            totalFound++;
            string referenced = vPattern.Match(strVal).Value.TrimEnd('\\');
            string refFolder = Path.GetFileName(referenced);
            string displayName = string.IsNullOrEmpty(valueName) ? "(Default)" : valueName;

            if (strVal.Contains(best.FolderName, StringComparison.OrdinalIgnoreCase))
            {
                Log($"  OK : HKCR\\{path}\\{displayName}", Color.DarkGreen);
                Log($"       -> {strVal}", Color.Gray);
            }
            else if (!Directory.Exists(referenced))
            {
                Log($"  INVALIDE : HKCR\\{path}\\{displayName}", Color.Red);
                Log($"       -> {strVal}", Color.Red);
                Log($"       ({refFolder} introuvable, sera corrigé vers {best.FolderName})", Color.FromArgb(200, 100, 0));
                issues++;
            }
            else
            {
                Log($"  ANCIEN : HKCR\\{path}\\{displayName}", Color.DarkOrange);
                Log($"       -> {strVal}", Color.DarkOrange);
                Log($"       ({refFolder} existe mais n'est pas la version la plus récente)", Color.Gray);
            }
        }

        // Recurse into DefaultIcon, shell\open\command, etc.
        foreach (string subName in key.GetSubKeyNames())
        {
            try
            {
                using var sub = key.OpenSubKey(subName);
                if (sub == null) continue;
                issues += ScanHKCRSubKey(sub, $"{path}\\{subName}", best, vPattern, ref totalFound);
            }
            catch { }
        }

        return issues;
    }

    private void BtnModifier_Click(object? sender, EventArgs e)
    {
        if (_bestInstallPath == null || _detectedVersions.Count == 0)
            return;

        var best = _detectedVersions.OrderByDescending(v => v.VersionNumber).First();

        var result = MessageBox.Show(
            $"Corriger tous les chemins d'icônes vers {best.FolderName} ?\n\n" +
            $"Chemin : {best.FullPath}\n\n" +
            "Les anciennes références V6xx introuvables seront remplacées.",
            "Confirmation",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

        if (result != DialogResult.Yes)
            return;

        btnModifier.Enabled = false;
        lblStatus.Text = "Correction en cours...";
        txtLog.Clear();

        int fixed_ = 0;
        int errors = 0;

        // Fix HKLM
        string[] registryRoots = new[]
        {
            @"SOFTWARE\Missler Software",
            @"SOFTWARE\WOW6432Node\Missler Software",
        };

        foreach (string root in registryRoots)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(root, writable: true);
                if (key == null) continue;
                FixKeyRecursive(key, root, best, ref fixed_, ref errors);
            }
            catch (Exception ex)
            {
                Log($"Erreur {root}: {ex.Message}", Color.Red);
                errors++;
            }
        }

        // Fix HKCR
        try
        {
            FixClassesRoot(best, ref fixed_, ref errors);
            FixFileAssociations(best, ref fixed_, ref errors);
        }
        catch (Exception ex)
        {
            Log($"Erreur HKCR: {ex.Message}", Color.Red);
            errors++;
        }

        Log("", Color.Black);
        Log($"Terminé : {fixed_} correction(s), {errors} erreur(s).", fixed_ > 0 ? Color.DarkGreen : Color.Black);
        lblStatus.Text = $"Terminé - {fixed_} correction(s)";

        if (fixed_ > 0)
        {
            // Notify shell of icon cache change
            NativeIconRefresh();
            Log("Cache d'icônes rafraîchi.", Color.DarkBlue);
        }

        btnModifier.Enabled = true;
    }

    private void FixKeyRecursive(RegistryKey parentKey, string parentPath, V6Installation best, ref int fixed_, ref int errors)
    {
        var vPattern = new Regex(@"C:\\Missler\\V\d{3}\\", RegexOptions.IgnoreCase);

        // Fix values at this level
        foreach (string valueName in parentKey.GetValueNames())
        {
            try
            {
                var val = parentKey.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (val is not string strVal) continue;

                if (vPattern.IsMatch(strVal) && !strVal.Contains(best.FolderName, StringComparison.OrdinalIgnoreCase))
                {
                    string referenced = vPattern.Match(strVal).Value.TrimEnd('\\');
                    if (!Directory.Exists(referenced))
                    {
                        string newVal = vPattern.Replace(strVal, $@"C:\Missler\{best.FolderName}\");
                        parentKey.SetValue(valueName, newVal);
                        Log($"  CORRIGÉ: {parentPath}\\{valueName}", Color.DarkGreen);
                        fixed_++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  ERREUR: {parentPath}\\{valueName}: {ex.Message}", Color.Red);
                errors++;
            }
        }

        // Recurse subkeys
        foreach (string subName in parentKey.GetSubKeyNames())
        {
            try
            {
                using var sub = parentKey.OpenSubKey(subName, writable: true);
                if (sub == null) continue;
                FixKeyRecursive(sub, $"{parentPath}\\{subName}", best, ref fixed_, ref errors);
            }
            catch { }
        }
    }

    private void FixClassesRoot(V6Installation best, ref int fixed_, ref int errors)
    {
        var vPattern = new Regex(@"C:\\Missler\\V\d{3}\\", RegexOptions.IgnoreCase);
        string[] prefixes = { "TopSolid", "Missler", ".top", ".tpd", ".tgp", ".tpr", ".tpp", ".dft", ".igs", ".stp", ".cam", ".mld", ".wod" };

        using var hkcr = Registry.ClassesRoot;
        foreach (string subName in hkcr.GetSubKeyNames())
        {
            bool relevant = false;
            foreach (var prefix in prefixes)
            {
                if (subName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                    subName.Equals(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    relevant = true;
                    break;
                }
            }
            if (!relevant) continue;

            try
            {
                using var sub = hkcr.OpenSubKey(subName, writable: true);
                if (sub == null) continue;
                FixHKCRSubKey(sub, subName, best, vPattern, ref fixed_, ref errors);
            }
            catch { }
        }
    }

    private readonly record struct FileAssociationRule(string Extension, string ProgId);

    private static readonly string[] LegacyAutoProgIds =
    {
        LegacyTopAutoProgId,
        LegacyDftAutoProgId,
    };

    private static readonly FileAssociationRule[] AssocRules =
    {
        new(@".top", PreferredTopProgId),
        new(@".dft", PreferredDftProgId),
    };

    private static string? FindTopSolidExecutable(string installPath, int versionNumber)
    {
        string[] preferredCandidates =
        {
            Path.Combine(installPath, "bin", $"top{versionNumber}.exe"),
            Path.Combine(installPath, "Top", "TopSolid.exe"),
            Path.Combine(installPath, "TopSolid.exe"),
        };

        foreach (string candidate in preferredCandidates)
        {
            if (File.Exists(candidate))
                return candidate;
        }

        string binPath = Path.Combine(installPath, "bin");
        if (Directory.Exists(binPath))
        {
            var topExe = Directory
                .GetFiles(binPath, "top*.exe", SearchOption.TopDirectoryOnly)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(topExe))
                return topExe;
        }

        return null;
    }

    private static string? GetTopSolidExePath(V6Installation best)
    {
        return FindTopSolidExecutable(best.FullPath, best.VersionNumber);
    }

    private static RegistryKey? OpenClassesScopeRoot(RegistryKey hiveRoot, string classesSubPath, bool writable)
    {
        return string.IsNullOrEmpty(classesSubPath)
            ? hiveRoot
            : hiveRoot.OpenSubKey(classesSubPath, writable);
    }

    private int AnalyzeFileAssociationIssues(V6Installation best)
    {
        int issues = 0;
        string? topExe = GetTopSolidExePath(best);

        var scopes = new[]
        {
            (Display: "HKCR", Root: Registry.ClassesRoot, ClassesPath: ""),
            (Display: "HKCU", Root: Registry.CurrentUser, ClassesPath: @"Software\Classes"),
            (Display: "HKLM", Root: Registry.LocalMachine, ClassesPath: @"SOFTWARE\Classes"),
        };

        foreach (var scope in scopes)
        {
            try
            {
                using var classesRoot = OpenClassesScopeRoot(scope.Root, scope.ClassesPath, writable: false);
                if (classesRoot == null) continue;

                foreach (var rule in AssocRules)
                {
                    using var extKey = classesRoot.OpenSubKey(rule.Extension);
                    string? actualProgId = extKey?.GetValue(null) as string;
                    bool mandatoryScope = string.Equals(scope.Display, "HKCR", StringComparison.OrdinalIgnoreCase);

                    if (extKey == null)
                    {
                        if (mandatoryScope)
                        {
                            Log($"  MANQUANT : {scope.Display}\\{rule.Extension}\\(Default)", Color.FromArgb(200, 100, 0));
                            Log($"       (sera défini vers {rule.ProgId})", Color.FromArgb(200, 100, 0));
                            issues++;
                        }
                    }
                    else if (!string.Equals(actualProgId, rule.ProgId, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"  INVALIDE : {scope.Display}\\{rule.Extension}\\(Default)", Color.Red);
                        Log($"       -> {actualProgId}", Color.Red);
                        Log($"       (sera corrigé vers {rule.ProgId})", Color.FromArgb(200, 100, 0));
                        issues++;
                    }

                    using var progIdKey = classesRoot.OpenSubKey(rule.ProgId);
                    if (progIdKey == null && !mandatoryScope)
                        continue;

                    using var commandKey = classesRoot.OpenSubKey($@"{rule.ProgId}\\shell\\open\\command");
                    string? command = commandKey?.GetValue(null) as string;
                    if (string.IsNullOrWhiteSpace(command))
                    {
                        if (mandatoryScope || progIdKey != null)
                        {
                            Log($"  MANQUANT : {scope.Display}\\{rule.ProgId}\\shell\\open\\command", Color.FromArgb(200, 100, 0));
                            issues++;
                        }
                    }
                    else if (topExe != null && !command.Contains(topExe, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"  INVALIDE : {scope.Display}\\{rule.ProgId}\\shell\\open\\command", Color.Red);
                        Log($"       -> {command}", Color.Red);
                        Log($"       (sera corrigé vers {topExe})", Color.FromArgb(200, 100, 0));
                        issues++;
                    }
                }

                foreach (string legacyProgId in LegacyAutoProgIds)
                {
                    using var legacyKey = classesRoot.OpenSubKey(legacyProgId);
                    if (legacyKey != null)
                    {
                        Log($"  ANCIEN : {scope.Display}\\{legacyProgId}", Color.DarkOrange);
                        Log("       (clé obsolète détectée, sera supprimée)", Color.Gray);
                        issues++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  Erreur analyse associations ({scope.Display}): {ex.Message}", Color.Red);
            }
        }

        if (topExe == null)
        {
            Log("  Erreur : exécutable TopSolid introuvable sur la version détectée.", Color.Red);
            issues++;
        }

        return issues;
    }

    private void FixFileAssociations(V6Installation best, ref int fixed_, ref int errors)
    {
        string? topExe = GetTopSolidExePath(best);
        if (topExe == null)
        {
            Log("  ERREUR: TopSolid.exe introuvable, correction associations impossible.", Color.Red);
            errors++;
            return;
        }

        string expectedCommand = $"\"{topExe}\" \"%1\"";
        const string expectedDde = "[open(\"%1\")]";

        var scopes = new[]
        {
            (Display: "HKCR", Root: Registry.ClassesRoot, ClassesPath: ""),
            (Display: "HKCU", Root: Registry.CurrentUser, ClassesPath: @"Software\Classes"),
            (Display: "HKLM", Root: Registry.LocalMachine, ClassesPath: @"SOFTWARE\Classes"),
        };

        foreach (var scope in scopes)
        {
            try
            {
                using var classesRoot = OpenClassesScopeRoot(scope.Root, scope.ClassesPath, writable: true);
                if (classesRoot == null) continue;

                foreach (var rule in AssocRules)
                {
                    bool mandatoryScope = string.Equals(scope.Display, "HKCR", StringComparison.OrdinalIgnoreCase);

                    RegistryKey? extKey = mandatoryScope
                        ? classesRoot.CreateSubKey(rule.Extension, writable: true)
                        : classesRoot.OpenSubKey(rule.Extension, writable: true);

                    if (extKey != null)
                    {
                        using (extKey)
                        {
                            string? actualProgId = extKey.GetValue(null) as string;
                            if (!string.Equals(actualProgId, rule.ProgId, StringComparison.OrdinalIgnoreCase))
                            {
                                extKey.SetValue(null, rule.ProgId);
                                Log($"  CORRIGÉ: {scope.Display}\\{rule.Extension}\\(Default) -> {rule.ProgId}", Color.DarkGreen);
                                fixed_++;
                            }
                        }
                    }

                    RegistryKey? progIdKey = mandatoryScope
                        ? classesRoot.CreateSubKey(rule.ProgId, writable: true)
                        : classesRoot.OpenSubKey(rule.ProgId, writable: true);

                    if (progIdKey != null)
                    {
                        using (progIdKey)
                        {
                            using var commandKey = progIdKey.CreateSubKey(@"shell\\open\\command", writable: true);
                            string? currentCommand = commandKey?.GetValue(null) as string;
                            if (commandKey != null && !string.Equals(currentCommand, expectedCommand, StringComparison.OrdinalIgnoreCase))
                            {
                                commandKey.SetValue(null, expectedCommand);
                                Log($"  CORRIGÉ: {scope.Display}\\{rule.ProgId}\\shell\\open\\command", Color.DarkGreen);
                                fixed_++;
                            }

                            using var ddeKey = progIdKey.CreateSubKey(@"shell\\open\\ddeexec", writable: true);
                            string? currentDde = ddeKey?.GetValue(null) as string;
                            if (ddeKey != null && !string.Equals(currentDde, expectedDde, StringComparison.OrdinalIgnoreCase))
                            {
                                ddeKey.SetValue(null, expectedDde);
                                Log($"  CORRIGÉ: {scope.Display}\\{rule.ProgId}\\shell\\open\\ddeexec", Color.DarkGreen);
                                fixed_++;
                            }
                        }
                    }
                }

                foreach (string legacyProgId in LegacyAutoProgIds)
                {
                    using var legacyKey = classesRoot.OpenSubKey(legacyProgId);
                    if (legacyKey != null)
                    {
                        classesRoot.DeleteSubKeyTree(legacyProgId, throwOnMissingSubKey: false);
                        Log($"  SUPPRIMÉ: {scope.Display}\\{legacyProgId}", Color.DarkGreen);
                        fixed_++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  ERREUR: correction associations ({scope.Display}): {ex.Message}", Color.Red);
                errors++;
            }
        }
    }

    private void FixHKCRSubKey(RegistryKey key, string path, V6Installation best, Regex vPattern, ref int fixed_, ref int errors)
    {
        foreach (string valueName in key.GetValueNames())
        {
            try
            {
                var val = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (val is not string strVal) continue;

                if (vPattern.IsMatch(strVal) && !strVal.Contains(best.FolderName, StringComparison.OrdinalIgnoreCase))
                {
                    string referenced = vPattern.Match(strVal).Value.TrimEnd('\\');
                    if (!Directory.Exists(referenced))
                    {
                        string newVal = vPattern.Replace(strVal, $@"C:\Missler\{best.FolderName}\");
                        key.SetValue(valueName, newVal);
                        Log($"  CORRIGÉ: HKCR\\{path}\\{valueName}", Color.DarkGreen);
                        fixed_++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"  ERREUR: HKCR\\{path}\\{valueName}: {ex.Message}", Color.Red);
                errors++;
            }
        }

        foreach (string subName in key.GetSubKeyNames())
        {
            try
            {
                using var sub = key.OpenSubKey(subName, writable: true);
                if (sub == null) continue;
                FixHKCRSubKey(sub, $"{path}\\{subName}", best, vPattern, ref fixed_, ref errors);
            }
            catch { }
        }
    }

    private void Log(string message, Color color)
    {
        txtLog.SelectionStart = txtLog.TextLength;
        txtLog.SelectionColor = color;
        txtLog.AppendText(message + Environment.NewLine);
        txtLog.ScrollToCaret();
    }

    private static void NativeIconRefresh()
    {
        // SHChangeNotify to refresh shell icon cache
        try
        {
            SHChangeNotify(0x08000000, 0x0000, IntPtr.Zero, IntPtr.Zero);
        }
        catch { }
    }

    [System.Runtime.InteropServices.DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
}
