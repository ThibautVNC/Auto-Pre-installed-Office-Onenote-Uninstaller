#Requires -RunAsAdministrator
<#
    .SYNOPSIS
        Removes pre-installed Microsoft 365 / OneNote (Click-to-Run) products.

    .DESCRIPTION
        Scans the uninstall registry hives for Click-to-Run Office packages,
        shows them in a menu and removes them silently with live progress.
        Interface available in English, Dutch and French.

    .NOTES
        Credit: Thibaut VNC
#>

$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------------ Strings --

$Strings = @{

    EN = @{
        WindowTitle   = 'Uninstall pre-installed Office  -  Thibaut VNC'
        BannerTitle   = '  UNINSTALL PRE-INSTALLED OFFICE / ONENOTE (Click-to-Run) '
        CreditLine    = '                                      Credit: Thibaut VNC '
        LangLabel     = 'Language: English'
        Found         = 'Found packages ({0}):'
        NoneFound     = 'No Microsoft 365 / OneNote Click-to-Run products found.'
        MenuStart     = '[1] Start removal'
        MenuRescan    = '[2] Rescan'
        MenuLang      = '[3] Language'
        MenuExit      = '[0] Exit'
        Choice        = '  Choice'
        Confirm       = '  Remove {0} package(s)? (Y/N)'
        YesPattern    = '^[YyJjOo]'
        NothingToDo   = 'Nothing to remove.'
        Cancelled     = 'Cancelled.'
        HeaderRemoval = '  REMOVAL STARTED'
        HeaderSummary = '  SUMMARY'
        Busy          = "`r        {0} removing...  {1:mm\:ss} elapsed "
        ParseFail     = 'UninstallString could not be parsed - skipped'
        Failed        = 'Failed: {0}'
        Duration      = '        {0}  (duration: {1:mm\:ss})'
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
    }

    NL = @{
        WindowTitle   = 'Uninstall pre-installed Office  -  Thibaut VNC'
        BannerTitle   = '  UNINSTALL PRE-INSTALLED OFFICE / ONENOTE (Click-to-Run) '
        CreditLine    = '                                      Credit: Thibaut VNC '
        LangLabel     = 'Taal: Nederlands'
        Found         = 'Gevonden pakketten ({0}):'
        NoneFound     = 'Geen Microsoft 365 / OneNote Click-to-Run producten gevonden.'
        MenuStart     = '[1] Verwijderen starten'
        MenuRescan    = '[2] Opnieuw scannen'
        MenuLang      = '[3] Taal'
        MenuExit      = '[0] Afsluiten'
        Choice        = '  Keuze'
        Confirm       = '  {0} pakket(ten) verwijderen? (J/N)'
        YesPattern    = '^[JjYy]'
        NothingToDo   = 'Niets te verwijderen.'
        Cancelled     = 'Geannuleerd.'
        HeaderRemoval = '  VERWIJDEREN GESTART'
        HeaderSummary = '  SAMENVATTING'
        Busy          = "`r        {0} bezig met verwijderen...  {1:mm\:ss} verstreken "
        ParseFail     = 'UninstallString kon niet gelezen worden - overgeslagen'
        Failed        = 'Mislukt: {0}'
        Duration      = '        {0}  (duur: {1:mm\:ss})'
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
    }

    FR = @{
        WindowTitle   = 'Uninstall pre-installed Office  -  Thibaut VNC'
        BannerTitle   = '  DESINSTALLER OFFICE / ONENOTE PREINSTALLE (Click-to-Run)'
        CreditLine    = '                                      Credit: Thibaut VNC '
        LangLabel     = 'Langue : Francais'
        Found         = 'Paquets trouves ({0}) :'
        NoneFound     = 'Aucun produit Microsoft 365 / OneNote Click-to-Run trouve.'
        MenuStart     = '[1] Demarrer la desinstallation'
        MenuRescan    = '[2] Analyser a nouveau'
        MenuLang      = '[3] Langue'
        MenuExit      = '[0] Quitter'
        Choice        = '  Choix'
        Confirm       = '  Supprimer {0} paquet(s) ? (O/N)'
        YesPattern    = '^[OoYyJj]'
        NothingToDo   = 'Rien a supprimer.'
        Cancelled     = 'Annule.'
        HeaderRemoval = '  DESINSTALLATION DEMARREE'
        HeaderSummary = '  RESUME'
        Busy          = "`r        {0} suppression en cours...  {1:mm\:ss} ecoulees "
        ParseFail     = 'UninstallString illisible - ignore'
        Failed        = 'Echec : {0}'
        Duration      = '        {0}  (duree : {1:mm\:ss})'
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
    }
}

# Default language
$Lang = 'EN'
$T    = $Strings[$Lang]

$Host.UI.RawUI.WindowTitle = $T.WindowTitle

# ---------------------------------------------------------------- Functions --

function Show-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  =========================================================' -ForegroundColor DarkCyan
    Write-Host $T.BannerTitle -ForegroundColor Cyan
    Write-Host '  =========================================================' -ForegroundColor DarkCyan
    Write-Host $T.CreditLine -ForegroundColor DarkGray
    Write-Host ''
}

function Get-OfficeTargets {
    Get-ChildItem -Path `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', `
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty |
        Where-Object { $_.DisplayName -match '^(Microsoft 365|Microsoft OneNote) - ' }
}

function Show-Targets {
    param([array]$Targets)

    if (-not $Targets) {
        Write-Host "  $($T.NoneFound)" -ForegroundColor Yellow
        Write-Host ''
        return
    }

    Write-Host ('  ' + ($T.Found -f $Targets.Count)) -ForegroundColor White
    $i = 1
    foreach ($t in $Targets) {
        Write-Host ('    {0}. {1}' -f $i, $t.DisplayName) -ForegroundColor Gray
        $i++
    }
    Write-Host ''
}

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

function Select-Language {
    Show-Banner
    Write-Host $T.LangHeader -ForegroundColor White
    Write-Host ''
    Write-Host '  [1] English'    -ForegroundColor Gray
    Write-Host '  [2] Nederlands' -ForegroundColor Gray
    Write-Host '  [3] Francais'   -ForegroundColor Gray
    Write-Host ''

    switch (Read-Host $T.Choice) {
        '1' { return 'EN' }
        '2' { return 'NL' }
        '3' { return 'FR' }
        default { return $script:Lang }
    }
}

function Start-Removal {
    param([array]$Targets)

    $Total        = $Targets.Count
    $Index        = 0
    $Results      = @()
    $RebootNeeded = $false

    Write-Host '  ---------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host $T.HeaderRemoval -ForegroundColor White
    Write-Host '  ---------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host ''

    foreach ($Target in $Targets) {
        $Index++
        $Name = $Target.DisplayName

        Write-Host "  [$Index/$Total] $Name" -ForegroundColor Cyan

        if ($Target.UninstallString -notmatch '^"([^"]+)"\s*(.*)$') {
            Write-Host "        $($T.ParseFail)" -ForegroundColor Yellow
            Write-Host ''
            $Results += [pscustomobject]@{ Name = $Name; Status = $T.StSkipped }
            continue
        }

        $Exe     = $Matches[1]
        $ArgList = "$($Matches[2]) displaylevel=False"

        try {
            $Watch   = [System.Diagnostics.Stopwatch]::StartNew()
            $Process = Start-Process -FilePath $Exe -ArgumentList $ArgList -PassThru -NoNewWindow

            # Live progress while the uninstaller runs
            $Frames = @('|', '/', '-', '\')
            $f = 0
            while (-not $Process.HasExited) {
                Write-Host ($T.Busy -f $Frames[$f % $Frames.Count], $Watch.Elapsed) `
                    -NoNewline -ForegroundColor DarkGray
                Start-Sleep -Milliseconds 250
                $f++
            }
            $Watch.Stop()
            Write-Host ("`r" + (' ' * 62) + "`r") -NoNewline   # clear the line

            $Code   = $Process.ExitCode
            $Status = Get-ExitCodeText -Code $Code
            $Colour = if ($Code -in 0, 1641, 3010) { 'Green' } else { 'Yellow' }
            if ($Code -in 1641, 3010) { $RebootNeeded = $true }

            Write-Host ($T.Duration -f $Status, $Watch.Elapsed) -ForegroundColor $Colour
            Write-Host ''

            $Results += [pscustomobject]@{ Name = $Name; Status = $Status }
        }
        catch {
            Write-Host ("`r" + (' ' * 62) + "`r") -NoNewline
            Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
            Write-Host ''
            $Results += [pscustomobject]@{ Name = $Name; Status = $T.StFailed }
        }
    }

    # ---- Summary ----
    Write-Host '  ---------------------------------------------------------' -ForegroundColor DarkGray
    Write-Host $T.HeaderSummary -ForegroundColor White
    Write-Host '  ---------------------------------------------------------' -ForegroundColor DarkGray
    foreach ($r in $Results) {
        Write-Host ('    {0,-40} {1}' -f $r.Name, $r.Status) -ForegroundColor Gray
    }
    Write-Host ''

    if ($RebootNeeded) {
        Write-Host "  $($T.RebootNote)" -ForegroundColor Yellow
        Write-Host ''
    }
}

# --------------------------------------------------------------------- Menu --

do {
    Show-Banner

    $Targets = @(Get-OfficeTargets)
    Show-Targets -Targets $Targets

    Write-Host "  $($T.MenuStart)"  -ForegroundColor White
    Write-Host "  $($T.MenuRescan)" -ForegroundColor White
    Write-Host "  $($T.MenuLang) - $($T.LangLabel)" -ForegroundColor White
    Write-Host "  $($T.MenuExit)"   -ForegroundColor White
    Write-Host ''
    $Choice = Read-Host $T.Choice

    switch ($Choice) {
        '1' {
            if (-not $Targets) {
                Write-Host ''
                Write-Host "  $($T.NothingToDo)" -ForegroundColor Yellow
            }
            else {
                Write-Host ''
                $Confirm = Read-Host ($T.Confirm -f $Targets.Count)
                if ($Confirm -match $T.YesPattern) {
                    Write-Host ''
                    Start-Removal -Targets $Targets
                }
                else {
                    Write-Host "  $($T.Cancelled)" -ForegroundColor Yellow
                }
            }
            Write-Host ''
            Read-Host $T.PressEnter | Out-Null
        }

        '2' { continue }

        '3' {
            $Lang = Select-Language
            $T    = $Strings[$Lang]
            $Host.UI.RawUI.WindowTitle = $T.WindowTitle
        }

        '0' {
            Write-Host ''
            Write-Host "  $($T.Closed)" -ForegroundColor DarkGray
            Write-Host ''
        }

        default {
            Write-Host ''
            Write-Host "  $($T.Invalid)" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
} while ($Choice -ne '0')
