# Top'ICO - TopSolid Icon Repair Tool v1.0.2

## Description
Top'ICO est un outil de réparation automatique des icônes de fichiers TopSolid (.top, .dft, .vdx, .prj) sur Windows 10 et 11.

## Problème résolu
Les icônes des fichiers TopSolid peuvent être corrompues ou afficher des icônes génériques au lieu des icônes appropriées de TopSolid.

## Fonctionnalités
- ✅ Détection automatique de l'installation TopSolid
- ✅ Réparation des associations d'icônes pour .top, .dft, .vdx, .prj
- ✅ Compatible Windows 10 et 11
- ✅ Fonctionne en mode utilisateur (pas besoin d'admin)
- ✅ Scripts Python et PowerShell inclus

## Installation
1. Téléchargez la dernière release depuis [GitHub Releases](https://github.com/3dainde/TopSolid-ICO/releases)
2. Extrayez l'archive
3. Lancez `fix_icons_simple.ps1` (PowerShell) ou `fix_icons.py` (Python)

## Utilisation

### PowerShell (recommandé)
```powershell
powershell -ExecutionPolicy Bypass -File fix_icons_simple.ps1
```

### Python
```bash
python fix_icons.py
```

### Interface graphique (optionnel)
Lancez `TopICO.exe` pour une interface graphique simple.

## Fichiers inclus
- `TopICO.exe` - Interface graphique (optionnel)
- `fix_icons.py` - Script Python de réparation
- `fix_icons_simple.ps1` - Script PowerShell simple
- `fix_icons.ps1` - Script PowerShell avancé
- `README.md` - Ce fichier

## Compatibilité
- ✅ Windows 10
- ✅ Windows 11
- ✅ TopSolid V6.24 à V6.27
- ✅ Python 3.8+ (pour le script Python)
- ✅ PowerShell 5.1+ (pour les scripts PowerShell)

## Historique des versions
- **v1.0.2** (2026-04-27): Scripts améliorés, compatibilité Windows 10/11, auto-détection TopSolid
- **v1.0.1**: Version initiale avec interface graphique

## Support
Si vous rencontrez des problèmes, vérifiez :
1. Que TopSolid est installé
2. Que vous avez les droits d'écriture sur le registre utilisateur
3. Que les fichiers TopSolid existent sur le Bureau

## Licence
MIT License - Voir LICENSE pour plus de détails.