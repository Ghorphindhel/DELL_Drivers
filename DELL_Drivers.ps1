#Requires -Version 5.1

<#
.SYNOPSIS
    Compare les pilotes de plusieurs packages Dell Latitude
    et génère un rapport Excel.

.PREREQUIS
    Les packages Dell doivent être extraits.
    Chaque répertoire doit contenir les fichiers .INF des pilotes.

    Module requis :
        Install-Module ImportExcel -Scope CurrentUser

.EXEMPLE D'ARBORESCENCE

    C:\DellDrivers\
        Latitude_5550\
        Latitude_5540\
        Latitude_5530\
        Latitude_5520\
#>

# -------------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------------

$Packages = [ordered]@{
    "Latitude 5550" = "C:\DellDrivers\Latitude_5550"
    "Latitude 5540" = "C:\DellDrivers\Latitude_5540"
    "Latitude 5530" = "C:\DellDrivers\Latitude_5530"
    "Latitude 5520" = "C:\DellDrivers\Latitude_5520"
}

$OutputFile = "C:\DellDrivers\Comparatif_Drivers_Latitude.xlsx"

# -------------------------------------------------------------------------
# VERIFICATION DU MODULE IMPORTEXCEL
# -------------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host ""
    Write-Host "Le module ImportExcel n'est pas installé." -ForegroundColor Yellow
    Write-Host "Installation :" -ForegroundColor Yellow
    Write-Host "Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}

Import-Module ImportExcel

# -------------------------------------------------------------------------
# FONCTIONS
# -------------------------------------------------------------------------

function Get-InfValue {
    param(
        [string[]]$Content,
        [string]$Property
    )

    $Line = $Content |
        Where-Object {
            $_ -match "^\s*$([regex]::Escape($Property))\s*="
        } |
        Select-Object -First 1

    if ($Line) {
        return (($Line -split "=", 2)[1]).Trim().Trim('"')
    }

    return ""
}


function Resolve-InfString {
    param(
        [string]$Value,
        [string[]]$Content
    )

    if (-not $Value) {
        return ""
    }

    # Exemple :
    # Provider = %Dell%
    # puis dans [Strings]
    # Dell = "Dell Technologies"

    if ($Value -match '^%(.+)%$') {

        $StringName = $Matches[1]

        $StringLine = $Content |
            Where-Object {
                $_ -match "^\s*$([regex]::Escape($StringName))\s*="
            } |
            Select-Object -First 1

        if ($StringLine) {
            return (($StringLine -split "=", 2)[1]).Trim().Trim('"')
        }
    }

    return $Value.Trim('"')
}


function Get-DriverFromInf {

    param(
        [string]$InfPath,
        [string]$Model,
        [string]$PackageRoot
    )

    try {
        $Content = Get-Content -LiteralPath $InfPath -ErrorAction Stop
    }
    catch {
        Write-Warning "Impossible de lire : $InfPath"
        return
    }

    # ---------------------------------------------------------------------
    # Provider
    # ---------------------------------------------------------------------

    $ProviderRaw = Get-InfValue -Content $Content -Property "Provider"
    $Provider = Resolve-InfString -Value $ProviderRaw -Content $Content

    # ---------------------------------------------------------------------
    # Class
    # ---------------------------------------------------------------------

    $Class = Get-InfValue -Content $Content -Property "Class"

    # ---------------------------------------------------------------------
    # ClassGUID
    # ---------------------------------------------------------------------

    $ClassGuid = Get-InfValue -Content $Content -Property "ClassGuid"

    # ---------------------------------------------------------------------
    # DriverVer
    #
    # Exemple :
    # DriverVer = 02/15/2025,32.0.101.6557
    # ---------------------------------------------------------------------

    $DriverVer = Get-InfValue -Content $Content -Property "DriverVer"

    $DriverDate = ""
    $DriverVersion = ""

    if ($DriverVer) {

        $Parts = $DriverVer -split ","

        if ($Parts.Count -ge 1) {
            $DriverDate = $Parts[0].Trim()
        }

        if ($Parts.Count -ge 2) {
            $DriverVersion = $Parts[1].Trim()
        }
    }

    # ---------------------------------------------------------------------
    # Catalog
    # ---------------------------------------------------------------------

    $Catalog = Get-InfValue -Content $Content -Property "CatalogFile"

    # ---------------------------------------------------------------------
    # Manufacturer
    # ---------------------------------------------------------------------

    $ManufacturerRaw = Get-InfValue -Content $Content -Property "Manufacturer"

    if ($ManufacturerRaw) {
        # On supprime ce qui suit une virgule
        $ManufacturerRaw = ($ManufacturerRaw -split ",")[0].Trim()
    }

    $Manufacturer = Resolve-InfString `
        -Value $ManufacturerRaw `
        -Content $Content

    # ---------------------------------------------------------------------
    # Chemin relatif
    # ---------------------------------------------------------------------

    try {
        $RelativePath = $InfPath.Substring($PackageRoot.Length).TrimStart("\")
    }
    catch {
        $RelativePath = $InfPath
    }

    # ---------------------------------------------------------------------
    # Nom du package / dossier parent
    # ---------------------------------------------------------------------

    $ParentFolder = Split-Path (Split-Path $InfPath -Parent) -Leaf

    # ---------------------------------------------------------------------
    # Identité du pilote
    #
    # On utilise INF + Provider + Class pour identifier une même famille
    # de pilote entre les différents modèles.
    # ---------------------------------------------------------------------

    $InfName = [System.IO.Path]::GetFileName($InfPath)

    $DriverKey = (
        $InfName.ToLower() + "|" +
        $Provider.ToLower() + "|" +
        $Class.ToLower()
    )

    [PSCustomObject]@{
        Modele       = $Model
        DriverKey    = $DriverKey
        INF           = $InfName
        Fournisseur   = $Provider
        Fabricant     = $Manufacturer
        Classe        = $Class
        ClassGUID     = $ClassGuid
        Version       = $DriverVersion
        Date          = $DriverDate
        Catalogue     = $Catalog
        Dossier       = $ParentFolder
        CheminRelatif = $RelativePath
        CheminComplet = $InfPath
    }
}


# -------------------------------------------------------------------------
# ANALYSE DES PACKAGES
# -------------------------------------------------------------------------

$AllDrivers = @()

foreach ($Model in $Packages.Keys) {

    $Path = $Packages[$Model]

    Write-Host ""
    Write-Host "Analyse : $Model" -ForegroundColor Cyan
    Write-Host "Chemin  : $Path"

    if (-not (Test-Path $Path)) {
        Write-Warning "Répertoire introuvable : $Path"
        continue
    }

    $InfFiles = Get-ChildItem `
        -Path $Path `
        -Filter "*.inf" `
        -File `
        -Recurse `
        -ErrorAction SilentlyContinue

    Write-Host "$($InfFiles.Count) fichiers INF trouvés."

    foreach ($Inf in $InfFiles) {

        $Driver = Get-DriverFromInf `
            -InfPath $Inf.FullName `
            -Model $Model `
            -PackageRoot $Path

        if ($Driver) {
            $AllDrivers += $Driver
        }
    }
}


# -------------------------------------------------------------------------
# SUPPRESSION DES DOUBLONS AU SEIN D'UN MEME MODELE
# -------------------------------------------------------------------------

$AllDrivers = $AllDrivers |
    Sort-Object Modele, DriverKey, Version -Unique


# -------------------------------------------------------------------------
# MODELES EFFECTIVEMENT ANALYSES
# -------------------------------------------------------------------------

$Models = @($Packages.Keys)
$ModelCount = $Models.Count


# -------------------------------------------------------------------------
# IDENTIFICATION DES PILOTES COMMUNS
# -------------------------------------------------------------------------

$Groups = $AllDrivers |
    Group-Object DriverKey


$CommonKeys = @()

foreach ($Group in $Groups) {

    $PresentModels = @(
        $Group.Group.Modele |
        Sort-Object -Unique
    )

    if ($PresentModels.Count -eq $ModelCount) {
        $CommonKeys += $Group.Name
    }
}


# -------------------------------------------------------------------------
# FEUILLE : COMMUNS AUX 4 MODELES
# -------------------------------------------------------------------------

$CommonDrivers = foreach ($Key in $CommonKeys) {

    $Drivers = $AllDrivers |
        Where-Object DriverKey -eq $Key

    $First = $Drivers | Select-Object -First 1

    $Versions = $Drivers.Version |
        Where-Object { $_ } |
        Sort-Object -Unique

    $SameVersion = if ($Versions.Count -eq 1) {
        "Oui"
    }
    else {
        "Non"
    }

    $Result = [ordered]@{
        INF              = $First.INF
        Fournisseur      = $First.Fournisseur
        Fabricant        = $First.Fabricant
        Classe           = $First.Classe
        "Meme version"   = $SameVersion
    }

    foreach ($Model in $Models) {

        $Driver = $Drivers |
            Where-Object Modele -eq $Model |
            Select-Object -First 1

        $Result["$Model - Version"] = $Driver.Version
        $Result["$Model - Date"]    = $Driver.Date
    }

    [PSCustomObject]$Result
}


# -------------------------------------------------------------------------
# FEUILLE : DIFFERENCES
# -------------------------------------------------------------------------

$Differences = foreach ($Group in $Groups) {

    $Drivers = $Group.Group
    $First = $Drivers | Select-Object -First 1

    $PresentModels = @(
        $Drivers.Modele |
        Sort-Object -Unique
    )

    $Versions = @(
        $Drivers.Version |
        Where-Object { $_ } |
        Sort-Object -Unique
    )

    $IsMissingModel = $PresentModels.Count -lt $ModelCount
    $DifferentVersion = $Versions.Count -gt 1

    if ($IsMissingModel -or $DifferentVersion) {

        $Reasons = @()

        if ($IsMissingModel) {
            $Reasons += "Absent sur certains modèles"
        }

        if ($DifferentVersion) {
            $Reasons += "Versions différentes"
        }

        $Result = [ordered]@{
            INF          = $First.INF
            Fournisseur  = $First.Fournisseur
            Fabricant    = $First.Fabricant
            Classe       = $First.Classe
            Difference   = ($Reasons -join " + ")
        }

        foreach ($Model in $Models) {

            $Driver = $Drivers |
                Where-Object Modele -eq $Model |
                Select-Object -First 1

            if ($Driver) {
                $Result[$Model] = $Driver.Version
            }
            else {
                $Result[$Model] = "ABSENT"
            }
        }

        [PSCustomObject]$Result
    }
}


# -------------------------------------------------------------------------
# FEUILLE : MATRICE
# -------------------------------------------------------------------------

$Matrix = foreach ($Group in $Groups) {

    $Drivers = $Group.Group
    $First = $Drivers | Select-Object -First 1

    $Result = [ordered]@{
        INF         = $First.INF
        Fournisseur = $First.Fournisseur
        Fabricant   = $First.Fabricant
        Classe      = $First.Classe
    }

    foreach ($Model in $Models) {

        $Driver = $Drivers |
            Where-Object Modele -eq $Model |
            Select-Object -First 1

        if ($Driver) {
            $Result[$Model] = $Driver.Version
        }
        else {
            $Result[$Model] = "ABSENT"
        }
    }

    [PSCustomObject]$Result
}


# -------------------------------------------------------------------------
# STATISTIQUES
# -------------------------------------------------------------------------

$Statistics = foreach ($Model in $Models) {

    $DriversForModel = $AllDrivers |
        Where-Object Modele -eq $Model

    [PSCustomObject]@{
        Modele               = $Model
        "Nombre de pilotes"  = $DriversForModel.Count
    }
}

$Statistics += [PSCustomObject]@{
    Modele              = "Communs aux 4 modèles"
    "Nombre de pilotes" = $CommonKeys.Count
}

$Statistics += [PSCustomObject]@{
    Modele              = "Pilotes avec différences"
    "Nombre de pilotes" = $Differences.Count
}


# -------------------------------------------------------------------------
# CREATION EXCEL
# -------------------------------------------------------------------------

if (Test-Path $OutputFile) {
    Remove-Item $OutputFile -Force
}

$ExcelParams = @{
    Path         = $OutputFile
    AutoSize     = $true
    AutoFilter   = $true
    FreezeTopRow = $true
    BoldTopRow   = $true
}

$Statistics |
    Export-Excel @ExcelParams `
    -WorksheetName "Resume"

$AllDrivers |
    Sort-Object Modele, Classe, Fournisseur, INF |
    Export-Excel @ExcelParams `
    -WorksheetName "Tous_les_pilotes" `
    -Append

$CommonDrivers |
    Sort-Object Classe, Fournisseur, INF |
    Export-Excel @ExcelParams `
    -WorksheetName "Communs_aux_4" `
    -Append

$Differences |
    Sort-Object Classe, Fournisseur, INF |
    Export-Excel @ExcelParams `
    -WorksheetName "Differences" `
    -Append

$Matrix |
    Sort-Object Classe, Fournisseur, INF |
    Export-Excel @ExcelParams `
    -WorksheetName "Matrice" `
    -Append


Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Analyse terminée." -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Fichier généré :" -ForegroundColor Cyan
Write-Host $OutputFile
Write-Host ""
Write-Host "Pilotes analysés : $($AllDrivers.Count)"
Write-Host "Pilotes communs   : $($CommonKeys.Count)"
Write-Host "Différences       : $($Differences.Count)"
Write-Host ""