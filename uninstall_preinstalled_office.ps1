#Requires -RunAsAdministrator
<#
    .SYNOPSIS
        Removes pre-installed Microsoft 365 / OneNote, Copilot and OneDrive.

    .DESCRIPTION
        Menu-driven cleanup tool for freshly imaged or OEM Windows machines.
        Select "Everything below" or pick individual components, review what
        was found, confirm, and let it run.
        Interface in English, Dutch and French.

    .NOTES
        Version : 1.3
        Credit  : Thibaut VNC
#>

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.3'

# ------------------------------------------------------------------ Strings --

$Strings = @{

    EN = @{
        WindowTitle   = 'Remove pre-installed Office / Copilot / OneDrive  -  Thibaut VNC'
        BannerTitle   = '   REMOVE PRE-INSTALLED OFFICE / COPILOT / ONEDRIVE'
        CreditLine    = '                       v{0}   -   Credit: Thibaut VNC'
        LangLabel     = 'English'
        SelectHeader  = 'Select what may be removed:'
        OptAll        = 'Everything below'
        OptOffice     = 'Microsoft 365 / OneNote'
        OptCopilot    = 'Copilot'
        OptOneDrive   = 'OneDrive'
        FoundCount    = '{0} found'
        NotFound      = 'not present'
        Locked        = 'locked'
        MenuStart     = '[1] Start removal'
        MenuRescan    = '[3] Rescan'
        MenuLang      = '[4] Language'
        MenuExit      = '[0] Exit'
        NavHint       = 'Up/Down to move, Space to tick'
        NavHintKeys   = 'Type A / O / C / D to tick, digits for the menu'
        Choice        = '  Choice'
        HintAll       = 'Deselect [1] to pick items individually.'
        HintItems     = 'Deselect all items to unlock [1] Everything below.'
        NothingSel    = 'Nothing selected.'
        NothingFound  = 'Nothing found to remove for your selection.'
        Confirm       = '  Remove the selected components? (Y/N)'
        YesPattern    = '^[YyJjOo]'
        Cancelled     = 'Cancelled.'
        WarnOneDrive  = 'Careful: OneDrive may be actively syncing company files on this machine.'
        HeaderRemoval = '  REMOVAL STARTED'
        HeaderSummary = '  SUMMARY'
        SecOffice     = '--- Microsoft 365 / OneNote ---'
        SecCopilot    = '--- Copilot ---'
        SecOneDrive   = '--- OneDrive ---'
        Busy          = "`r        {0} working...  {1:mm\:ss} elapsed "
        ParseFail     = 'UninstallString could not be parsed - skipped'
        Failed        = 'Failed: {0}'
        Duration      = '        {0}  (duration: {1:mm\:ss})'
        NoneHere      = '        Nothing to do.'
        RebootNote    = 'Note: a restart is required to finish the cleanup.'
        PressEnter    = '  Press Enter to return to the menu'
        Invalid       = 'Invalid choice.'
        Closed        = 'Closed.  -  Credit: Thibaut VNC'
        LangHeader    = '  Select a language:'
        StOk          = 'OK - removed successfully'
        StReboot      = 'OK - restart required'
        StUserCancel  = 'Cancelled by user'
        StInterrupted = 'Interrupted (scenario not completed)'
        StCode        = 'Exited with code {0}'
        StSkipped     = 'Skipped'
        StFailed      = 'Failed'
        StRemoved     = 'Removed'
        StProtected   = 'Protected by Windows - left in place'
        StPolicySet   = 'Policy set (stays disabled)'
        Scanning      = 'Scanning...'
        ScanOffice    = 'Office / OneNote'
        ScanCopilot   = 'Copilot (this takes a few seconds)'
        ScanOneDrive  = 'OneDrive'
        MenuPick      = '[2] Choose which Office items'
        PickHeader    = 'Which Office / OneNote items may be removed?'
        PickHint      = 'Untick a language to keep it installed.'
        PickAll       = '[A] Tick all'
        PickNone      = '[N] Tick none'
        PickBack      = '[0] Back'
        PickHint2     = 'Up/Down to move, Space to tick'
        SelCount      = '({0} ticked)'
        NoneTicked    = 'No Office items are ticked.'
    }

    NL = @{
        WindowTitle   = 'Voorgeinstalleerde Office / Copilot / OneDrive wissen  -  Thibaut VNC'
        BannerTitle   = '   VOORGEINSTALLEERDE OFFICE / COPILOT / ONEDRIVE WISSEN'
        CreditLine    = '                       v{0}   -   Credit: Thibaut VNC'
        LangLabel     = 'Nederlands'
        SelectHeader  = 'Selecteer wat verwijderd mag worden:'
        OptAll        = 'Alles hieronder'
        OptOffice     = 'Microsoft 365 / OneNote'
        OptCopilot    = 'Copilot'
        OptOneDrive   = 'OneDrive'
        FoundCount    = '{0} gevonden'
        NotFound      = 'niet aanwezig'
        Locked        = 'vergrendeld'
        MenuStart     = '[1] Verwijderen starten'
        MenuRescan    = '[3] Opnieuw scannen'
        MenuLang      = '[4] Taal'
        MenuExit      = '[0] Afsluiten'
        NavHint       = 'Pijltjes om te navigeren, spatie om aan te vinken'
        NavHintKeys   = 'Typ A / O / C / D om aan te vinken, cijfers voor het menu'
        Choice        = '  Keuze'
        HintAll       = 'Deselecteer [1] om items apart te kiezen.'
        HintItems     = 'Deselecteer alle items om [1] Alles hieronder vrij te geven.'
        NothingSel    = 'Niets geselecteerd.'
        NothingFound  = 'Niets gevonden om te verwijderen voor je selectie.'
        Confirm       = '  Geselecteerde onderdelen verwijderen? (J/N)'
        YesPattern    = '^[JjYy]'
        Cancelled     = 'Geannuleerd.'
        WarnOneDrive  = 'Opgelet: OneDrive synchroniseert op dit toestel mogelijk bedrijfsbestanden.'
        HeaderRemoval = '  VERWIJDEREN GESTART'
        HeaderSummary = '  SAMENVATTING'
        SecOffice     = '--- Microsoft 365 / OneNote ---'
        SecCopilot    = '--- Copilot ---'
        SecOneDrive   = '--- OneDrive ---'
        Busy          = "`r        {0} bezig...  {1:mm\:ss} verstreken "
        ParseFail     = 'UninstallString kon niet gelezen worden - overgeslagen'
        Failed        = 'Mislukt: {0}'
        Duration      = '        {0}  (duur: {1:mm\:ss})'
        NoneHere      = '        Niets te doen.'
        RebootNote    = 'Let op: een herstart is vereist om het opruimen af te ronden.'
        PressEnter    = '  Druk op Enter om terug te keren naar het menu'
        Invalid       = 'Ongeldige keuze.'
        Closed        = 'Afgesloten.  -  Credit: Thibaut VNC'
        LangHeader    = '  Kies een taal:'
        StOk          = 'OK - succesvol verwijderd'
        StReboot      = 'OK - herstart vereist'
        StUserCancel  = 'Geannuleerd door gebruiker'
        StInterrupted = 'Onderbroken (scenario niet voltooid)'
        StCode        = 'Afgesloten met code {0}'
        StSkipped     = 'Overgeslagen'
        StFailed      = 'Mislukt'
        StRemoved     = 'Verwijderd'
        StProtected   = 'Beschermd door Windows - niet verwijderd'
        StPolicySet   = 'Policy ingesteld (blijft uitgeschakeld)'
        Scanning      = 'Scannen...'
        ScanOffice    = 'Office / OneNote'
        ScanCopilot   = 'Copilot (dit duurt enkele seconden)'
        ScanOneDrive  = 'OneDrive'
        MenuPick      = '[2] Kies welke Office-items'
        PickHeader    = 'Welke Office / OneNote items mogen weg?'
        PickHint      = 'Vink een taal uit om die te behouden.'
        PickAll       = '[A] Alles aanvinken'
        PickNone      = '[N] Alles uitvinken'
        PickBack      = '[0] Terug'
        PickHint2     = 'Pijltjes om te navigeren, spatie om aan te vinken'
        SelCount      = '({0} aangevinkt)'
        NoneTicked    = 'Er zijn geen Office-items aangevinkt.'
    }

    FR = @{
        WindowTitle   = 'Supprimer Office / Copilot / OneDrive preinstalles  -  Thibaut VNC'
        BannerTitle   = '   SUPPRIMER OFFICE / COPILOT / ONEDRIVE PREINSTALLES'
        CreditLine    = '                       v{0}   -   Credit: Thibaut VNC'
        LangLabel     = 'Francais'
        SelectHeader  = 'Selectionnez ce qui peut etre supprime :'
        OptAll        = 'Tout ci-dessous'
        OptOffice     = 'Microsoft 365 / OneNote'
        OptCopilot    = 'Copilot'
        OptOneDrive   = 'OneDrive'
        FoundCount    = '{0} trouve(s)'
        NotFound      = 'absent'
        Locked        = 'verrouille'
        MenuStart     = '[1] Demarrer la suppression'
        MenuRescan    = '[3] Analyser a nouveau'
        MenuLang      = '[4] Langue'
        MenuExit      = '[0] Quitter'
        NavHint       = 'Fleches pour naviguer, Espace pour cocher'
        NavHintKeys   = 'Tapez A / O / C / D pour cocher, chiffres pour le menu'
        Choice        = '  Choix'
        HintAll       = 'Deselectionnez [1] pour choisir les elements un par un.'
        HintItems     = 'Deselectionnez tous les elements pour deverrouiller [1].'
        NothingSel    = 'Rien de selectionne.'
        NothingFound  = 'Rien a supprimer pour votre selection.'
        Confirm       = '  Supprimer les composants selectionnes ? (O/N)'
        YesPattern    = '^[OoYyJj]'
        Cancelled     = 'Annule.'
        WarnOneDrive  = 'Attention : OneDrive synchronise peut-etre des fichiers d entreprise ici.'
        HeaderRemoval = '  SUPPRESSION DEMARREE'
        HeaderSummary = '  RESUME'
        SecOffice     = '--- Microsoft 365 / OneNote ---'
        SecCopilot    = '--- Copilot ---'
        SecOneDrive   = '--- OneDrive ---'
        Busy          = "`r        {0} en cours...  {1:mm\:ss} ecoulees "
        ParseFail     = 'UninstallString illisible - ignore'
        Failed        = 'Echec : {0}'
        Duration      = '        {0}  (duree : {1:mm\:ss})'
        NoneHere      = '        Rien a faire.'
        RebootNote    = 'Attention : un redemarrage est requis pour terminer le nettoyage.'
        PressEnter    = '  Appuyez sur Entree pour revenir au menu'
        Invalid       = 'Choix invalide.'
        Closed        = 'Ferme.  -  Credit: Thibaut VNC'
        LangHeader    = '  Choisissez une langue :'
        StOk          = 'OK - supprime avec succes'
        StReboot      = 'OK - redemarrage requis'
        StUserCancel  = 'Annule par l utilisateur'
        StInterrupted = 'Interrompu (scenario non termine)'
        StCode        = 'Termine avec le code {0}'
        StSkipped     = 'Ignore'
        StFailed      = 'Echec'
        StRemoved     = 'Supprime'
        StProtected   = 'Protege par Windows - conserve'
        StPolicySet   = 'Strategie appliquee (reste desactive)'
        Scanning      = 'Analyse en cours...'
        ScanOffice    = 'Office / OneNote'
        ScanCopilot   = 'Copilot (quelques secondes)'
        ScanOneDrive  = 'OneDrive'
        MenuPick      = '[2] Choisir les elements Office'
        PickHeader    = 'Quels elements Office / OneNote peuvent etre supprimes ?'
        PickHint      = 'Decochez une langue pour la conserver.'
        PickAll       = '[A] Tout cocher'
        PickNone      = '[N] Tout decocher'
        PickBack      = '[0] Retour'
        PickHint2     = 'Fleches pour naviguer, Espace pour cocher'
        SelCount      = '({0} coche(s))'
        NoneTicked    = 'Aucun element Office n est coche.'
    }
}

$Lang = 'EN'
$T    = $Strings[$Lang]

$Host.UI.RawUI.WindowTitle = $T.WindowTitle

# Selection state
$Sel = @{ All = $false; Office = $false; Copilot = $false; OneDrive = $false }

$Script:Results      = @()
$Script:RebootNeeded = $false

# --------------------------------------------------------------- Detection --

function Get-OfficeTargets {
    Get-ChildItem -Path `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', `
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty |
        Where-Object { $_.DisplayName -match '^(Microsoft 365|Microsoft OneNote) - ' }
}

function Get-CopilotTargets {
    $Found = @()

    # -Name filters inside the Appx API, which is far faster than pulling
    # every package into the pipeline and filtering with Where-Object.
    $Found += Get-AppxPackage -AllUsers -Name '*Copilot*' -ErrorAction SilentlyContinue |
              ForEach-Object {
                  [pscustomobject]@{ Kind = 'Appx'; Name = $_.Name; Id = $_.PackageFullName }
              }

    $Found += Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
              Where-Object { $_.DisplayName -match 'Copilot' } |
              ForEach-Object {
                  [pscustomobject]@{ Kind = 'Provisioned'; Name = $_.DisplayName; Id = $_.PackageName }
              }

    $Found
}

function Invoke-Scan {
    Write-Host "  $($T.Scanning)" -ForegroundColor White

    Write-Host "    - $($T.ScanOffice)" -NoNewline -ForegroundColor DarkGray
    $Script:Office = @(Get-OfficeTargets)

    # Every detected Office item is ticked by default; the picker lets you
    # untick individual language SKUs you want to keep.
    $Script:OfficePick = @()
    foreach ($o in $Script:Office) { $Script:OfficePick += $true }

    Write-Host '  ok' -ForegroundColor DarkGray

    Write-Host "    - $($T.ScanCopilot)" -NoNewline -ForegroundColor DarkGray
    $Script:Copilot = @(Get-CopilotTargets)
    Write-Host '  ok' -ForegroundColor DarkGray

    Write-Host "    - $($T.ScanOneDrive)" -NoNewline -ForegroundColor DarkGray
    $Script:OneDrive = @(Get-OneDriveTargets)
    Write-Host '  ok' -ForegroundColor DarkGray
}

function Get-OneDriveTargets {
    $Paths = @(
        "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
        "$env:SystemRoot\System32\OneDriveSetup.exe"
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe"
    )

    $Found = @()
    foreach ($p in $Paths) {
        if (Test-Path $p) {
            $Found += [pscustomobject]@{ Name = 'Microsoft OneDrive'; Path = $p }
        }
    }
    $Found
}

# ---------------------------------------------------------------- Helpers ---

function Get-ExitCodeText {
    param([int]$Code)

    switch ($Code) {
        0     { $T.StOk }
        1641  { $T.StReboot }
        3010  { $T.StReboot }
        1602  { $T.StUserCancel }
        17002 { $T.StInterrupted }
        default { $T.StCode -f $Code }
    }
}

function Wait-WithProgress {
    param([System.Diagnostics.Process]$Process)

    $Watch  = [System.Diagnostics.Stopwatch]::StartNew()
    $Frames = @('|', '/', '-', '\')
    $f = 0

    while (-not $Process.HasExited) {
        Write-Host ($T.Busy -f $Frames[$f % $Frames.Count], $Watch.Elapsed) `
            -NoNewline -ForegroundColor DarkGray
        Start-Sleep -Milliseconds 250
        $f++
    }
    $Watch.Stop()
    Write-Host ("`r" + (' ' * 64) + "`r") -NoNewline
    return $Watch.Elapsed
}

function Add-Result {
    param([string]$Component, [string]$Name, [string]$Status)
    $Script:Results += [pscustomobject]@{
        Component = $Component
        Name      = $Name
        Status    = $Status
    }
}

# ----------------------------------------------------------------- Removal --

function Remove-OfficeTargets {
    param([array]$Targets)

    Write-Host "  $($T.SecOffice)" -ForegroundColor White

    if (-not $Targets) {
        Write-Host $T.NoneHere -ForegroundColor DarkGray
        Write-Host ''
        return
    }

    $Total = $Targets.Count
    $Index = 0

    foreach ($Target in $Targets) {
        $Index++
        $Name = $Target.DisplayName
        Write-Host "  [$Index/$Total] $Name" -ForegroundColor Cyan

        if ($Target.UninstallString -notmatch '^"([^"]+)"\s*(.*)$') {
            Write-Host "        $($T.ParseFail)" -ForegroundColor Yellow
            Add-Result 'Office' $Name $T.StSkipped
            continue
        }

        $Exe     = $Matches[1]
        $ArgList = "$($Matches[2]) displaylevel=False"

        try {
            $Process = Start-Process -FilePath $Exe -ArgumentList $ArgList -PassThru -NoNewWindow
            $Elapsed = Wait-WithProgress -Process $Process

            $Code   = $Process.ExitCode
            $Status = Get-ExitCodeText -Code $Code
            $Colour = if ($Code -in 0, 1641, 3010) { 'Green' } else { 'Yellow' }
            if ($Code -in 1641, 3010) { $Script:RebootNeeded = $true }

            Write-Host ($T.Duration -f $Status, $Elapsed) -ForegroundColor $Colour
            Add-Result 'Office' $Name $Status
        }
        catch {
            Write-Host ("`r" + (' ' * 64) + "`r") -NoNewline
            Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
            Add-Result 'Office' $Name $T.StFailed
        }
    }
    Write-Host ''
}

function Remove-CopilotTargets {
    param([array]$Targets)

    Write-Host "  $($T.SecCopilot)" -ForegroundColor White

    if (-not $Targets) {
        Write-Host $T.NoneHere -ForegroundColor DarkGray
    }
    else {
        $Total = $Targets.Count
        $Index = 0

        foreach ($Target in $Targets) {
            $Index++
            Write-Host "  [$Index/$Total] $($Target.Name)  ($($Target.Kind))" -ForegroundColor Cyan

            try {
                if ($Target.Kind -eq 'Appx') {
                    Remove-AppxPackage -Package $Target.Id -AllUsers -ErrorAction Stop
                }
                else {
                    Remove-AppxProvisionedPackage -Online -PackageName $Target.Id -ErrorAction Stop | Out-Null
                }
                Write-Host "        $($T.StRemoved)" -ForegroundColor Green
                Add-Result 'Copilot' $Target.Name $T.StRemoved
            }
            catch {
                # Some Copilot components are system apps and cannot be uninstalled
                if ($_.Exception.Message -match '0x80073CFA|not authorized|system app') {
                    Write-Host "        $($T.StProtected)" -ForegroundColor Yellow
                    Add-Result 'Copilot' $Target.Name $T.StProtected
                }
                else {
                    Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
                    Add-Result 'Copilot' $Target.Name $T.StFailed
                }
            }
        }
    }

    # Policy so Copilot does not come back after a feature update
    try {
        $Key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot'
        if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }
        New-ItemProperty -Path $Key -Name 'TurnOffWindowsCopilot' -Value 1 -PropertyType DWord -Force | Out-Null
        Write-Host "        $($T.StPolicySet)" -ForegroundColor Green
        Add-Result 'Copilot' 'TurnOffWindowsCopilot' $T.StPolicySet
    }
    catch {
        Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
    }

    Write-Host ''
}

function Remove-OneDriveTargets {
    param([array]$Targets)

    Write-Host "  $($T.SecOneDrive)" -ForegroundColor White

    if (-not $Targets) {
        Write-Host $T.NoneHere -ForegroundColor DarkGray
        Write-Host ''
        return
    }

    Get-Process -Name 'OneDrive' -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue

    $Total = $Targets.Count
    $Index = 0

    foreach ($Target in $Targets) {
        $Index++
        Write-Host "  [$Index/$Total] $($Target.Path)" -ForegroundColor Cyan

        try {
            $Process = Start-Process -FilePath $Target.Path -ArgumentList '/uninstall' -PassThru -NoNewWindow
            $Elapsed = Wait-WithProgress -Process $Process

            $Code   = $Process.ExitCode
            $Status = Get-ExitCodeText -Code $Code
            $Colour = if ($Code -in 0, 1641, 3010) { 'Green' } else { 'Yellow' }
            if ($Code -in 1641, 3010) { $Script:RebootNeeded = $true }

            Write-Host ($T.Duration -f $Status, $Elapsed) -ForegroundColor $Colour
            Add-Result 'OneDrive' (Split-Path $Target.Path -Leaf) $Status
        }
        catch {
            Write-Host ("`r" + (' ' * 64) + "`r") -NoNewline
            Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
            Add-Result 'OneDrive' (Split-Path $Target.Path -Leaf) $T.StFailed
        }
    }
    Write-Host ''
}

# --------------------------------------------------------------------- UI ---

# Some hosts (PowerShell ISE, a few remoting scenarios) cannot read single
# keypresses. Detected on first use, with a typed fallback so the tool still
# works everywhere.
$Script:RawKeys = $true

function Read-MenuKey {
    if ($Script:RawKeys) {
        try {
            $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

            switch ($Key.VirtualKeyCode) {
                38 { return 'UP' }
                40 { return 'DOWN' }
                32 { return 'SPACE' }
                13 { return 'ENTER' }
                27 { return 'ESC' }
            }

            $Char = $Key.Character
            if ($Char -match '[0-9A-Za-z]') { return ([string]$Char).ToUpper() }
            return ''
        }
        catch {
            $Script:RawKeys = $false
        }
    }

    return (Read-Host $T.Choice).Trim().ToUpper()
}

function Show-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host $T.BannerTitle -ForegroundColor Cyan
    Write-Host '  ============================================================' -ForegroundColor DarkCyan
    Write-Host ($T.CreditLine -f $ScriptVersion) -ForegroundColor DarkGray
    Write-Host ''
}

function Show-Row {
    param(
        [bool]$IsCursor,
        [bool]$Checked,
        [bool]$Available,
        [string]$Label,
        [string]$Detail
    )

    $Arrow = if ($IsCursor) { '>' } else { ' ' }

    if (-not $Available) {
        Write-Host ('   {0}  {1}  {2,-26} {3}' -f $Arrow, '[-]', $Label, "($($T.Locked))") `
            -ForegroundColor DarkGray
        return
    }

    $Box = if ($Checked) { '[x]' } else { '[ ]' }

    if ($IsCursor) {
        $Colour = 'Cyan'
    }
    elseif ($Checked) {
        $Colour = 'Green'
    }
    else {
        $Colour = 'Gray'
    }

    Write-Host ('   {0}  {1}  {2,-26} {3}' -f $Arrow, $Box, $Label, $Detail) -ForegroundColor $Colour
}

function Get-RowAvailable {
    param([int]$Row)

    $ItemsSelected = $Sel.Office -or $Sel.Copilot -or $Sel.OneDrive

    if ($Row -eq 0) { return (-not $ItemsSelected) }
    return (-not $Sel.All)
}

function Move-Cursor {
    param([int]$Current, [int]$Delta)

    $Row = $Current
    for ($n = 0; $n -lt 4; $n++) {
        $Row = $Row + $Delta
        if ($Row -lt 0) { $Row = 3 }
        if ($Row -gt 3) { $Row = 0 }
        if (Get-RowAvailable -Row $Row) { return $Row }
    }
    return $Current
}

function Switch-Row {
    param([int]$Row)

    if (-not (Get-RowAvailable -Row $Row)) { return }

    switch ($Row) {
        0 { $Sel.All      = -not $Sel.All }
        1 { $Sel.Office   = -not $Sel.Office }
        2 { $Sel.Copilot  = -not $Sel.Copilot }
        3 { $Sel.OneDrive = -not $Sel.OneDrive }
    }
}

function Select-Language {
    Show-Banner
    Write-Host $T.LangHeader -ForegroundColor White
    Write-Host ''
    Write-Host '   [1] English'    -ForegroundColor Gray
    Write-Host '   [2] Nederlands' -ForegroundColor Gray
    Write-Host '   [3] Francais'   -ForegroundColor Gray
    Write-Host '   [0] ...'        -ForegroundColor DarkGray
    Write-Host ''

    switch (Read-MenuKey) {
        '1' { return 'EN' }
        '2' { return 'NL' }
        '3' { return 'FR' }
        default { return $script:Lang }
    }
}

function Get-OfficePicked {
    $Picked = @()
    for ($i = 0; $i -lt $Script:Office.Count; $i++) {
        if ($Script:OfficePick[$i]) { $Picked += $Script:Office[$i] }
    }
    $Picked
}

function Select-OfficeItems {
    $Cursor = 0

    do {
        Show-Banner
        Write-Host "  $($T.PickHeader)" -ForegroundColor White
        Write-Host "  $($T.PickHint)" -ForegroundColor DarkGray
        Write-Host ''

        for ($i = 0; $i -lt $Script:Office.Count; $i++) {
            $Arrow = if ($i -eq $Cursor) { '>' } else { ' ' }
            $Box   = if ($Script:OfficePick[$i]) { '[x]' } else { '[ ]' }

            if ($i -eq $Cursor) {
                $Colour = 'Cyan'
            }
            elseif ($Script:OfficePick[$i]) {
                $Colour = 'Green'
            }
            else {
                $Colour = 'DarkGray'
            }

            Write-Host ('   {0}  {1}  {2}' -f $Arrow, $Box, $Script:Office[$i].DisplayName) `
                -ForegroundColor $Colour
        }

        Write-Host ''
        if ($Script:RawKeys) {
            Write-Host "   $($T.PickHint2)" -ForegroundColor DarkGray
        }
        Write-Host ''
        Write-Host "  $($T.PickAll)"  -ForegroundColor White
        Write-Host "  $($T.PickNone)" -ForegroundColor White
        Write-Host "  $($T.PickBack)" -ForegroundColor White
        Write-Host ''

        $Pick = Read-MenuKey

        switch ($Pick) {
            'UP'    { $Cursor--; if ($Cursor -lt 0) { $Cursor = $Script:Office.Count - 1 } }
            'DOWN'  { $Cursor++; if ($Cursor -ge $Script:Office.Count) { $Cursor = 0 } }
            'SPACE' { $Script:OfficePick[$Cursor] = -not $Script:OfficePick[$Cursor] }
            'A'     { for ($i = 0; $i -lt $Script:OfficePick.Count; $i++) { $Script:OfficePick[$i] = $true } }
            'N'     { for ($i = 0; $i -lt $Script:OfficePick.Count; $i++) { $Script:OfficePick[$i] = $false } }
            'ESC'   { $Pick = '0' }
            'ENTER' { $Pick = '0' }
            default {
                # Typed fallback: a row number toggles that row
                if ($Pick -match '^\d+$' -and $Pick -ne '0') {
                    $Idx = [int]$Pick - 1
                    if ($Idx -ge 0 -and $Idx -lt $Script:OfficePick.Count) {
                        $Script:OfficePick[$Idx] = -not $Script:OfficePick[$Idx]
                    }
                }
            }
        }
    } while ($Pick -ne '0')
}

function Get-Detail {
    param($Items)
    if ($Items.Count -gt 0) { $T.FoundCount -f $Items.Count } else { $T.NotFound }
}

function Get-OfficeDetail {
    if ($Script:Office.Count -eq 0) { return $T.NotFound }

    $Ticked = @(Get-OfficePicked).Count
    $Text   = $T.FoundCount -f $Script:Office.Count
    if ($Ticked -ne $Script:Office.Count) {
        $Text = "$Text  " + ($T.SelCount -f $Ticked)
    }
    $Text
}

# ------------------------------------------------------------------- Main ---

Show-Banner
Invoke-Scan

$Cursor = 0

do {
    $ItemsSelected = $Sel.Office -or $Sel.Copilot -or $Sel.OneDrive
    $AllAvailable  = -not $ItemsSelected
    $ItemAvailable = -not $Sel.All

    # Keep the cursor on a row that is actually selectable
    if (-not (Get-RowAvailable -Row $Cursor)) { $Cursor = Move-Cursor -Current $Cursor -Delta 1 }

    Show-Banner
    Write-Host "  $($T.SelectHeader)" -ForegroundColor White
    Write-Host ''

    Show-Row -IsCursor ($Cursor -eq 0) -Checked $Sel.All -Available $AllAvailable -Label $T.OptAll -Detail ''
    Write-Host '      ------------------------------------------------------' -ForegroundColor DarkGray
    Show-Row -IsCursor ($Cursor -eq 1) -Checked $Sel.Office   -Available $ItemAvailable -Label $T.OptOffice   -Detail (Get-OfficeDetail)
    Show-Row -IsCursor ($Cursor -eq 2) -Checked $Sel.Copilot  -Available $ItemAvailable -Label $T.OptCopilot  -Detail (Get-Detail $Copilot)
    Show-Row -IsCursor ($Cursor -eq 3) -Checked $Sel.OneDrive -Available $ItemAvailable -Label $T.OptOneDrive -Detail (Get-Detail $OneDrive)
    Write-Host ''

    if ($Script:RawKeys) {
        Write-Host "   $($T.NavHint)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "   $($T.NavHintKeys)" -ForegroundColor DarkGray
    }

    if ($Sel.All)           { Write-Host "   $($T.HintAll)"   -ForegroundColor DarkGray }
    elseif ($ItemsSelected) { Write-Host "   $($T.HintItems)" -ForegroundColor DarkGray }
    Write-Host ''

    Write-Host "  $($T.MenuStart)"  -ForegroundColor White
    if ($Office.Count -gt 1) { Write-Host "  $($T.MenuPick)" -ForegroundColor White }
    Write-Host "  $($T.MenuRescan)" -ForegroundColor White
    Write-Host "  $($T.MenuLang) - $($T.LangLabel)" -ForegroundColor White
    Write-Host "  $($T.MenuExit)"   -ForegroundColor White
    Write-Host ''

    $Choice = Read-MenuKey

    switch ($Choice) {

        'UP'    { $Cursor = Move-Cursor -Current $Cursor -Delta -1 }
        'DOWN'  { $Cursor = Move-Cursor -Current $Cursor -Delta  1 }
        'SPACE' { Switch-Row -Row $Cursor }

        # Letter shortcuts, mainly for the typed fallback
        'A' { Switch-Row -Row 0 }
        'O' { Switch-Row -Row 1 }
        'C' { Switch-Row -Row 2 }
        'D' { Switch-Row -Row 3 }

        '2' {
            if ($Office.Count -gt 1) { Select-OfficeItems }
        }

        '3' {
            Show-Banner
            Invoke-Scan
        }

        '4' {
            $Lang = Select-Language
            $T    = $Strings[$Lang]
            $Host.UI.RawUI.WindowTitle = $T.WindowTitle
        }

        { $_ -in '1', 'ENTER' } {

            $DoOffice   = $Sel.All -or $Sel.Office
            $DoCopilot  = $Sel.All -or $Sel.Copilot
            $DoOneDrive = $Sel.All -or $Sel.OneDrive

            if (-not ($DoOffice -or $DoCopilot -or $DoOneDrive)) {
                Write-Host ''
                Write-Host "  $($T.NothingSel)" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            $OfficeToRemove = @(Get-OfficePicked)

            if ($DoOffice -and $Office.Count -gt 0 -and $OfficeToRemove.Count -eq 0) {
                Write-Host ''
                Write-Host "  $($T.NoneTicked)" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            $Planned = 0
            if ($DoOffice)   { $Planned += $OfficeToRemove.Count }
            if ($DoCopilot)  { $Planned += $Copilot.Count }
            if ($DoOneDrive) { $Planned += $OneDrive.Count }

            if ($Planned -eq 0) {
                Write-Host ''
                Write-Host "  $($T.NothingFound)" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            Write-Host ''
            if ($DoOneDrive -and $OneDrive.Count -gt 0) {
                Write-Host "  $($T.WarnOneDrive)" -ForegroundColor Yellow
                Write-Host ''
            }

            $Confirm = Read-Host $T.Confirm
            if ($Confirm -notmatch $T.YesPattern) {
                Write-Host "  $($T.Cancelled)" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            $Script:Results      = @()
            $Script:RebootNeeded = $false

            Write-Host ''
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host $T.HeaderRemoval -ForegroundColor White
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host ''

            if ($DoOffice)   { Remove-OfficeTargets   -Targets $OfficeToRemove }
            if ($DoCopilot)  { Remove-CopilotTargets  -Targets $Copilot }
            if ($DoOneDrive) { Remove-OneDriveTargets -Targets $OneDrive }

            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host $T.HeaderSummary -ForegroundColor White
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            foreach ($r in $Script:Results) {
                Write-Host ('    {0,-9} {1,-32} {2}' -f $r.Component, $r.Name, $r.Status) -ForegroundColor Gray
            }
            Write-Host ''

            if ($Script:RebootNeeded) {
                Write-Host "  $($T.RebootNote)" -ForegroundColor Yellow
                Write-Host ''
            }

            # Refresh detection so the menu reflects the new state
            Invoke-Scan
            Write-Host ''

            Read-Host $T.PressEnter | Out-Null
        }

        '0' {
            Write-Host ''
            Write-Host "  $($T.Closed)" -ForegroundColor DarkGray
            Write-Host ''
        }
    }
} while ($Choice -ne '0')
