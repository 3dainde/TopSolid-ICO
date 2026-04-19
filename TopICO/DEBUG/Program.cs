// BreakIcons - DEBUG TOOL
// Casse volontairement les chemins d'icônes TopSolid dans le registre
// pour tester la réparation par Top'ICO.
//
// Remplace les chemins C:\Missler\V6xx\ par C:\Missler\V600\ (inexistant)
// ce qui rend les icônes blanches/génériques dans l'explorateur.

using Microsoft.Win32;
using System.Text.RegularExpressions;

const string FakePath = @"C:\Missler\V600\";
var vPattern = new Regex(@"C:\\Missler\\V6\d{2}\\", RegexOptions.IgnoreCase);

Console.ForegroundColor = ConsoleColor.Red;
Console.WriteLine("╔══════════════════════════════════════════════════╗");
Console.WriteLine("║  BreakIcons - DEBUG TopSolid Icon Paths          ║");
Console.WriteLine("║  Remplace tous les chemins V6xx par V600 (faux)  ║");
Console.WriteLine("╚══════════════════════════════════════════════════╝");
Console.ResetColor();
Console.WriteLine();
Console.WriteLine($"Les chemins C:\\Missler\\V6xx\\ seront remplacés par {FakePath}");
Console.WriteLine("Les icônes deviendront blanches/génériques.");
Console.WriteLine();
Console.ForegroundColor = ConsoleColor.Yellow;
Console.Write("Continuer ? (O/N) : ");
Console.ResetColor();

var answer = Console.ReadLine()?.Trim().ToUpperInvariant();
if (answer != "O" && answer != "OUI" && answer != "Y")
{
    Console.WriteLine("Annulé.");
    return;
}

Console.WriteLine();

int modified = 0;
int errors = 0;

// --- HKLM ---
string[] hklmRoots =
{
    @"SOFTWARE\Missler Software",
    @"SOFTWARE\WOW6432Node\Missler Software",
};

foreach (string root in hklmRoots)
{
    try
    {
        using var key = Registry.LocalMachine.OpenSubKey(root, writable: true);
        if (key == null)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine($"  (clé HKLM\\{root} absente, ignorée)");
            Console.ResetColor();
            continue;
        }
        BreakKeyRecursive(key, $"HKLM\\{root}", vPattern, ref modified, ref errors);
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"  ERREUR HKLM\\{root}: {ex.Message}");
        Console.ResetColor();
        errors++;
    }
}

// --- HKCR ---
string[] prefixes = { "TopSolid", "Missler", ".top", ".tpd", ".tgp", ".tpr", ".tpp", ".dft", ".igs", ".stp", ".cam", ".mld", ".wod" };

try
{
    using var hkcr = Registry.ClassesRoot;
    foreach (string subName in hkcr.GetSubKeyNames())
    {
        bool relevant = prefixes.Any(p =>
            subName.StartsWith(p, StringComparison.OrdinalIgnoreCase) ||
            subName.Equals(p, StringComparison.OrdinalIgnoreCase));
        if (!relevant) continue;

        try
        {
            using var sub = hkcr.OpenSubKey(subName, writable: true);
            if (sub == null) continue;
            BreakKeyRecursive(sub, $"HKCR\\{subName}", vPattern, ref modified, ref errors);
        }
        catch { }
    }
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"  ERREUR HKCR: {ex.Message}");
    Console.ResetColor();
    errors++;
}

// --- Rafraîchir le cache d'icônes ---
if (modified > 0)
{
    SHChangeNotify(0x08000000, 0x0000, nint.Zero, nint.Zero);
    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("\nCache d'icônes rafraîchi.");
    Console.ResetColor();
}

// --- Résumé ---
Console.WriteLine();
Console.ForegroundColor = modified > 0 ? ConsoleColor.Red : ConsoleColor.Green;
Console.WriteLine($"Terminé : {modified} valeur(s) cassée(s), {errors} erreur(s).");
Console.ResetColor();

if (modified > 0)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.WriteLine("\nLes icônes TopSolid devraient maintenant apparaître blanches.");
    Console.WriteLine("Lancez Top'ICO pour les réparer.");
    Console.ResetColor();
}

Console.WriteLine("\nAppuyez sur une touche pour fermer...");
Console.ReadKey(true);

// ─── Fonctions ───────────────────────────────────────────────

void BreakKeyRecursive(RegistryKey parentKey, string parentPath, Regex pattern, ref int mod, ref int err)
{
    foreach (string valueName in parentKey.GetValueNames())
    {
        try
        {
            var val = parentKey.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            if (val is not string strVal) continue;
            if (!pattern.IsMatch(strVal)) continue;

            // Déjà cassé vers V600 ?
            if (strVal.Contains(FakePath, StringComparison.OrdinalIgnoreCase)) continue;

            string newVal = pattern.Replace(strVal, FakePath);
            parentKey.SetValue(valueName, newVal);

            string displayName = string.IsNullOrEmpty(valueName) ? "(Default)" : valueName;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("  CASSÉ : ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"{parentPath}\\{displayName}");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine($"          {strVal}");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"       -> {newVal}");
            Console.ResetColor();
            mod++;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"  ERREUR: {parentPath}\\{valueName}: {ex.Message}");
            Console.ResetColor();
            err++;
        }
    }

    foreach (string subName in parentKey.GetSubKeyNames())
    {
        try
        {
            using var sub = parentKey.OpenSubKey(subName, writable: true);
            if (sub == null) continue;
            BreakKeyRecursive(sub, $"{parentPath}\\{subName}", pattern, ref mod, ref err);
        }
        catch { }
    }
}

[System.Runtime.InteropServices.DllImport("shell32.dll")]
static extern void SHChangeNotify(int wEventId, int uFlags, nint dwItem1, nint dwItem2);
