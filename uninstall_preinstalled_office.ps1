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
        Version : 1.9
        Credit  : Thibaut VNC
#>

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.9'

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
        MenuRescan    = '[2] Rescan'
        MenuLang      = '[3] Language'
        MenuReboot    = '[4] Reboot now'
        MenuExit      = '[0] Exit'
        ConfirmReboot = '  Reboot this machine now? (Y/N)'
        Rebooting     = 'Rebooting in 5 seconds - close this window to cancel.'
        RebootOffer   = '  A restart is needed. Reboot now? (Y/N)'
        ShellBack     = 'Windows Explorer was restarted (HP uninstallers close it).'
        WarnShell     = 'Heads up: the desktop and taskbar will disappear during this step. That is normal - Explorer is restarted afterwards.'
        RebootAsk     = '  Reboot now to finish up? (Y/N)'
        RunDone       = 'Run finished.'
        NavHint       = 'Up/Down move   Right/Enter unfold a group   Left fold   Space ticks'
        TagKeep       = '[not bloatware - kept]'
        NavHintKeys   = 'Type A / O / C / D / H to tick a group, digits for the menu'
        Choice        = '  Choice'
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
        SelCount      = '({0} ticked)'
        NoneTicked    = 'Nothing is ticked. Open a list with the right arrow and pick what may go.'
        OptHp         = 'HP bloatware'
        ScanHp        = 'HP software'
        SecHp         = '--- HP bloatware ---'
        EscSkip       = 'Esc skips this wait'
        TagLicensed   = '[licensed]'
        WarnLicensed  = 'WARNING: a LICENSED Microsoft 365 install is ticked for removal:'
        WarnLicensed2 = 'This is not OEM bloatware. Make sure you can reinstall and reactivate it.'
        StSkippedWait = 'Still running - skipped, check it after the reboot'
        StTimedOut    = 'Timed out - still running in the background'
        WarnHp        = 'Careful: HP Wolf Security is security software. Check company policy first.'
        StManual      = 'No silent uninstall - remove manually'
        StStillThere  = 'Reported OK but still installed - reboot and retry'
        Verifying     = 'verifying...'
        HpStuckNote   = 'Some HP components survived. Reboot, then run this again - HP protects them while they run.'
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
        MenuRescan    = '[2] Opnieuw scannen'
        MenuLang      = '[3] Taal'
        MenuReboot    = '[4] Nu herstarten'
        MenuExit      = '[0] Afsluiten'
        ConfirmReboot = '  Dit toestel nu herstarten? (J/N)'
        Rebooting     = 'Herstart over 5 seconden - sluit dit venster om te annuleren.'
        RebootOffer   = '  Een herstart is nodig. Nu herstarten? (J/N)'
        ShellBack     = 'Windows Verkenner is opnieuw gestart (HP-uninstallers sluiten die af).'
        WarnShell     = 'Let op: bureaublad en taakbalk verdwijnen tijdens deze stap. Dat is normaal - Verkenner wordt daarna opnieuw gestart.'
        RebootAsk     = '  Nu herstarten om af te ronden? (J/N)'
        RunDone       = 'Run afgerond.'
        NavHint       = 'Op/Neer navigeren   Rechts/Enter openklappen   Links dichtklappen   Spatie vinkt aan'
        TagKeep       = '[geen bloatware - blijft]'
        NavHintKeys   = 'Typ A / O / C / D / H voor een groep, cijfers voor het menu'
        Choice        = '  Keuze'
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
        SelCount      = '({0} aangevinkt)'
        NoneTicked    = 'Niets aangevinkt. Open een lijst met pijltje rechts en kies wat weg mag.'
        OptHp         = 'HP bloatware'
        ScanHp        = 'HP software'
        SecHp         = '--- HP bloatware ---'
        EscSkip       = 'Esc slaat dit wachten over'
        TagLicensed   = '[licensed]'
        WarnLicensed  = 'LET OP: een LICENSED Microsoft 365-installatie staat aangevinkt:'
        WarnLicensed2 = 'Dit is geen OEM-bloatware. Zorg dat je ze opnieuw kan installeren en activeren.'
        StSkippedWait = 'Draait nog - overgeslagen, check na de herstart'
        StTimedOut    = 'Time-out - draait nog op de achtergrond'
        WarnHp        = 'Opgelet: HP Wolf Security is beveiligingssoftware. Check eerst het bedrijfsbeleid.'
        StManual      = 'Geen stille uninstall - manueel verwijderen'
        StStillThere  = 'Meldde OK maar staat er nog - herstart en run opnieuw'
        Verifying     = 'controleren...'
        HpStuckNote   = 'Sommige HP-onderdelen staan er nog. Herstart en run opnieuw - HP beschermt ze zolang ze draaien.'
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
        MenuRescan    = '[2] Analyser a nouveau'
        MenuLang      = '[3] Langue'
        MenuReboot    = '[4] Redemarrer maintenant'
        MenuExit      = '[0] Quitter'
        ConfirmReboot = '  Redemarrer cette machine maintenant ? (O/N)'
        Rebooting     = 'Redemarrage dans 5 secondes - fermez cette fenetre pour annuler.'
        RebootOffer   = '  Un redemarrage est necessaire. Redemarrer maintenant ? (O/N)'
        ShellBack     = 'L Explorateur Windows a ete relance (les desinstalleurs HP le ferment).'
        WarnShell     = 'Attention : le bureau et la barre des taches vont disparaitre. C est normal - l Explorateur sera relance.'
        RebootAsk     = '  Redemarrer maintenant pour terminer ? (O/N)'
        RunDone       = 'Execution terminee.'
        NavHint       = 'Haut/Bas naviguer   Droite/Entree deplier   Gauche replier   Espace coche'
        TagKeep       = '[pas du bloatware - conserve]'
        NavHintKeys   = 'Tapez A / O / C / D / H pour un groupe, chiffres pour le menu'
        Choice        = '  Choix'
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
        SelCount      = '({0} coche(s))'
        NoneTicked    = 'Rien n est coche. Ouvrez une liste avec la fleche droite et choisissez.'
        OptHp         = 'Bloatware HP'
        ScanHp        = 'Logiciels HP'
        SecHp         = '--- Bloatware HP ---'
        EscSkip       = 'Echap passe cette attente'
        TagLicensed   = '[sous licence]'
        WarnLicensed  = 'ATTENTION : une installation Microsoft 365 SOUS LICENCE est cochee :'
        WarnLicensed2 = 'Ce n est pas du bloatware OEM. Assurez-vous de pouvoir la reinstaller et reactiver.'
        StSkippedWait = 'Toujours en cours - ignore, verifiez apres le redemarrage'
        StTimedOut    = 'Delai depasse - toujours en cours en arriere-plan'
        WarnHp        = 'Attention : HP Wolf Security est un logiciel de securite. Verifiez la politique.'
        StManual      = 'Pas de desinstallation silencieuse - a faire manuellement'
        StStillThere  = 'OK signale mais toujours installe - redemarrez et relancez'
        Verifying     = 'verification...'
        HpStuckNote   = 'Certains composants HP subsistent. Redemarrez puis relancez.'
    }
}

$Lang = 'EN'
$T    = $Strings[$Lang]

$Host.UI.RawUI.WindowTitle = $T.WindowTitle

# Selection state
$Sel = @{ All = $false; Office = $false; Copilot = $false; OneDrive = $false; Hp = $false }

# Which component groups are unfolded in the tree
$Script:Expanded = @{ 1 = $false; 2 = $false; 3 = $false; 4 = $false }

$Script:Results      = @()
$Script:RebootNeeded = $false
$Script:HpStuck      = $false
$Script:WaitAborted  = $false

# --------------------------------------------------------------- Detection --

function Get-OfficeTargets {
    # Display names are localised, so matching on English words like
    # "Apps for enterprise" misses most of them: the same product reads
    # "Microsoft 365-apps voor ondernemingen", "para Grandes Empresas",
    # "-appar for foretag"... What every Click-to-Run product DOES share is
    # an UninstallString that runs OfficeClickToRun.exe, so that is the gate.
    $OemPattern = '^(Microsoft 365|Microsoft OneNote) - '

    Get-ChildItem -Path `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', `
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty |
        Where-Object {
            $_.DisplayName -match $OemPattern -or
            ($_.DisplayName -match '^Microsoft (365|OneNote)' -and
             $_.UninstallString -match 'OfficeClickToRun')
        } |
        ForEach-Object {
            # Anything that is not the bare OEM "Microsoft 365 - xx-xx" form is
            # treated as a licensed install: detected and listed, never ticked
            # by default.
            $_ | Add-Member -NotePropertyName 'Licensed' `
                            -NotePropertyValue (-not ($_.DisplayName -match $OemPattern)) `
                            -PassThru -Force
        } |
        Sort-Object Licensed, DisplayName
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
    # Everything found starts ticked, licensed installs included: ticking the
    # row means "all of it". Licensed items are tagged in the item view so they
    # can be unticked one by one, and they raise a warning before anything runs.
    $Script:OfficePick = @()
    foreach ($o in $Script:Office) { $Script:OfficePick += $true }

    Write-Host '  ok' -ForegroundColor DarkGray

    Write-Host "    - $($T.ScanCopilot)" -NoNewline -ForegroundColor DarkGray
    $Script:Copilot = @(Get-CopilotTargets)

    $Script:CopilotPick = @()
    foreach ($c in $Script:Copilot) { $Script:CopilotPick += $true }

    Write-Host '  ok' -ForegroundColor DarkGray

    Write-Host "    - $($T.ScanOneDrive)" -NoNewline -ForegroundColor DarkGray
    $Script:OneDrive = @(Get-OneDriveTargets)

    $Script:OneDrivePick = @()
    foreach ($o in $Script:OneDrive) { $Script:OneDrivePick += $true }

    Write-Host '  ok' -ForegroundColor DarkGray

    Write-Host "    - $($T.ScanHp)" -NoNewline -ForegroundColor DarkGray
    $Script:Hp = @(Get-HpTargets)

    $Script:HpPick = @()
    foreach ($h in $Script:Hp) { $Script:HpPick += (-not $h.Keep) }

    Write-Host '  ok' -ForegroundColor DarkGray
}

function Get-HpTargets {
    # Curated list. Same principle as the Office pattern: match only what is
    # genuinely OEM bloatware, never a broad "HP*" sweep. Drivers, firmware,
    # HP Hotkey Support and audio components are deliberately absent, because
    # removing those breaks hardware function keys and sound.
    $Patterns = @(
        '^HP Wolf Security'
        '^HP Security Update Service'
        '^HP One Agent'
        '^HP Insights'
        '^HP Analytics'
        '^HP Sure (Click|Sense|Run|Recover|Start|Backup)'
        '^HP Client Security Manager'
        '^HP Support Assistant'
        '^HP JumpStart'
        '^HP Connection Optimizer'
        '^HP Documentation'
        '^HP Privacy Settings'
        '^HP System Information'
        '^HP Notifications'
        '^HP Desktop Support Utilities'
        '^HP QuickDrop'
        '^HP Registration Service'
        '^HP PC Hardware Diagnostics'
        '^myHP'
    )
    $Combined = ($Patterns -join '|')

    $Found = @()

    $Found += Get-ChildItem -Path `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', `
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty |
        Where-Object { $_.DisplayName -match $Combined } |
        ForEach-Object {
            [pscustomobject]@{
                Kind        = 'Registry'
                DisplayName = $_.DisplayName
                Quiet       = $_.QuietUninstallString
                Uninstall   = $_.UninstallString
                Id          = $null
                # Support Assistant delivers driver and BIOS updates, so it is
                # listed but never swept up by ticking the group.
                Keep        = [bool]($_.DisplayName -match '^HP Support Assistant')
            }
        }

    # HP Store apps all ship under the same publisher prefix
    $Found += Get-AppxPackage -AllUsers -Name 'AD2F1837.*' -ErrorAction SilentlyContinue |
              ForEach-Object {
                  [pscustomobject]@{
                      Kind        = 'Appx'
                      DisplayName = $_.Name
                      Quiet       = $null
                      Uninstall   = $null
                      Id          = $_.PackageFullName
                      Keep        = $false
                  }
              }

    # Wolf Security only uninstalls cleanly in this order
    $Weight = {
        param($Name)
        if ($Name -match '^HP Wolf Security - Console') { return 2 }
        if ($Name -match '^HP Wolf Security')           { return 1 }
        if ($Name -match '^HP Security Update Service') { return 3 }
        # One Agent re-deploys HP software, so it goes last
        if ($Name -match '^HP One Agent')               { return 20 }
        return 10
    }

    $Found | Sort-Object Keep, @{ Expression = { & $Weight $_.DisplayName } }, DisplayName
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

function Disable-QuickEdit {
    # With QuickEdit on (the Windows default), a stray click puts the console
    # into selection mode and the NEXT Write-Host blocks until a key is
    # pressed. The script freezes mid-timer and "resumes" when you hit Enter -
    # it was never stuck, just waiting on the console. Turning QuickEdit off
    # for this console removes the trap entirely.
    try {
        if (-not ('Win32.ConsoleMode' -as [type])) {
            Add-Type -Namespace Win32 -Name ConsoleMode -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true)]
public static extern IntPtr GetStdHandle(int nStdHandle);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
[DllImport("kernel32.dll", SetLastError = true)]
public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
'@
        }

        $Handle = [Win32.ConsoleMode]::GetStdHandle(-10)   # STD_INPUT_HANDLE
        $Mode   = 0
        if (-not [Win32.ConsoleMode]::GetConsoleMode($Handle, [ref]$Mode)) { return }

        $New = $Mode -bor 0x0080                            # ENABLE_EXTENDED_FLAGS
        if ($New -band 0x0040) { $New = $New - 0x0040 }     # clear ENABLE_QUICK_EDIT_MODE

        [void][Win32.ConsoleMode]::SetConsoleMode($Handle, [uint32]$New)
    }
    catch {
        # Not a real console (ISE, redirected output) - nothing to disarm
    }
}

function Test-EscPressed {
    # Non-blocking: lets a long wait be abandoned without killing the script
    try {
        while ([Console]::KeyAvailable) {
            if ([Console]::ReadKey($true).Key -eq 'Escape') { return $true }
        }
    }
    catch { }
    return $false
}

function Wait-WithProgress {
    param(
        [System.Diagnostics.Process]$Process,
        [int]$TimeoutMinutes = 120
    )

    $Script:WaitAborted = $false

    $Watch  = [System.Diagnostics.Stopwatch]::StartNew()
    $Frames = @('|', '/', '-', '\')
    $f = 0

    while (-not $Process.HasExited) {
        if (Test-EscPressed) {
            $Script:WaitAborted = 'skipped'
            break
        }
        if ($Watch.Elapsed.TotalMinutes -ge $TimeoutMinutes) {
            $Script:WaitAborted = 'timeout'
            break
        }

        Write-Host ($T.Busy -f $Frames[$f % $Frames.Count], $Watch.Elapsed) `
            -NoNewline -ForegroundColor DarkGray
        Start-Sleep -Milliseconds 250
        $f++
    }
    $Watch.Stop()
    Write-Host ("`r" + (' ' * 70) + "`r") -NoNewline
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

            if ($Script:WaitAborted) {
                $Status = if ($Script:WaitAborted -eq 'timeout') { $T.StTimedOut } else { $T.StSkippedWait }
                $Script:RebootNeeded = $true
                Write-Host "        $Status" -ForegroundColor Yellow
                Add-Result 'Office' $Name $Status
                continue
            }

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

function Invoke-UninstallCommand {
    param([string]$Command)

    if ($Command -match '^"([^"]+)"\s*(.*)$') {
        $Exe = $Matches[1]; $Arg = $Matches[2]
    }
    elseif ($Command -match '^(\S+)\s*(.*)$') {
        $Exe = $Matches[1]; $Arg = $Matches[2]
    }
    else {
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($Arg)) {
        return Start-Process -FilePath $Exe -PassThru -NoNewWindow
    }
    return Start-Process -FilePath $Exe -ArgumentList $Arg -PassThru -NoNewWindow
}

function Get-SilentCommand {
    param($Target)

    if ($Target.Quiet) { return $Target.Quiet }

    $U = $Target.Uninstall
    if (-not $U) { return $null }

    # MSI packages can always be made silent
    if ($U -match 'msiexec' -and $U -match '(\{[0-9A-Fa-f\-]{36}\})') {
        return "msiexec.exe /X$($Matches[1]) /qn /norestart"
    }

    # Anything else would pop a GUI and hang the run, so leave it to the user
    return $null
}

function Test-StillInstalled {
    param([string]$DisplayName)

    $Hit = Get-ChildItem -Path `
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', `
        'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' `
        -ErrorAction SilentlyContinue |
        Get-ItemProperty |
        Where-Object { $_.DisplayName -eq $DisplayName }

    return [bool]$Hit
}

function Wait-UntilGone {
    # Several HP uninstallers return exit code 0 within seconds while the real
    # work happens in a child process, so the exit code alone proves nothing.
    # Wait for the registry entry to actually disappear instead.
    param(
        [string]$DisplayName,
        [int]$TimeoutSeconds = 180
    )

    $Watch  = [System.Diagnostics.Stopwatch]::StartNew()
    $Frames = @('|', '/', '-', '\')
    $f = 0

    while ($Watch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (-not (Test-StillInstalled -DisplayName $DisplayName)) {
            $Watch.Stop()
            Write-Host ("`r" + (' ' * 70) + "`r") -NoNewline
            return $true
        }

        if (Test-EscPressed) { break }

        Write-Host ("`r        {0} {1}  {2:mm\:ss}   ({3}) " -f `
            $Frames[$f % $Frames.Count], $T.Verifying, $Watch.Elapsed, $T.EscSkip) `
            -NoNewline -ForegroundColor DarkGray

        # One second between polls: each check walks both uninstall hives
        Start-Sleep -Milliseconds 1000
        $f++
    }

    $Watch.Stop()
    Write-Host ("`r" + (' ' * 70) + "`r") -NoNewline
    return $false
}

function Restore-Shell {
    # HP uninstallers unload shell extensions, which takes explorer.exe with
    # them. Without this the desktop and taskbar stay gone until a reboot.
    if (-not (Get-Process -Name 'explorer' -ErrorAction SilentlyContinue)) {
        try {
            Start-Process -FilePath 'explorer.exe' -ErrorAction Stop
            Start-Sleep -Seconds 2
            Write-Host "        $($T.ShellBack)" -ForegroundColor DarkGray
        }
        catch { }
    }
}

function Remove-HpTargets {
    param([array]$Targets)

    Write-Host "  $($T.SecHp)" -ForegroundColor White

    if ($Targets) {
        Write-Host "  $($T.WarnShell)" -ForegroundColor Yellow
        Write-Host ''
    }

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

        try {
            if ($Target.Kind -eq 'Appx') {
                Remove-AppxPackage -Package $Target.Id -AllUsers -ErrorAction Stop
                Write-Host "        $($T.StRemoved)" -ForegroundColor Green
                Add-Result 'HP' $Name $T.StRemoved
                continue
            }

            $Command = Get-SilentCommand -Target $Target
            if (-not $Command) {
                Write-Host "        $($T.StManual)" -ForegroundColor Yellow
                Add-Result 'HP' $Name $T.StManual
                continue
            }

            $Process = Invoke-UninstallCommand -Command $Command
            if (-not $Process) {
                Write-Host "        $($T.ParseFail)" -ForegroundColor Yellow
                Add-Result 'HP' $Name $T.StSkipped
                continue
            }

            $Elapsed = Wait-WithProgress -Process $Process

            if ($Script:WaitAborted) {
                $Status = if ($Script:WaitAborted -eq 'timeout') { $T.StTimedOut } else { $T.StSkippedWait }
                $Script:HpStuck      = $true
                $Script:RebootNeeded = $true
                Write-Host "        $Status" -ForegroundColor Yellow
                Add-Result 'HP' $Name $Status
                continue
            }

            $Code = $Process.ExitCode
            if ($Code -in 1641, 3010) { $Script:RebootNeeded = $true }

            # Trust the registry, not the exit code
            $Gone = Wait-UntilGone -DisplayName $Name

            if ($Gone) {
                $Status = Get-ExitCodeText -Code $Code
                Write-Host ($T.Duration -f $Status, $Elapsed) -ForegroundColor Green
                Add-Result 'HP' $Name $Status
            }
            else {
                $Script:HpStuck      = $true
                $Script:RebootNeeded = $true
                Write-Host "        $($T.StStillThere)" -ForegroundColor Yellow
                Add-Result 'HP' $Name $T.StStillThere
            }
        }
        catch {
            Write-Host ("`r" + (' ' * 64) + "`r") -NoNewline
            Write-Host ('        ' + ($T.Failed -f $_.Exception.Message)) -ForegroundColor Red
            Add-Result 'HP' $Name $T.StFailed
        }
    }

    Restore-Shell
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

            if ($Script:WaitAborted) {
                $Status = if ($Script:WaitAborted -eq 'timeout') { $T.StTimedOut } else { $T.StSkippedWait }
                $Script:RebootNeeded = $true
                Write-Host "        $Status" -ForegroundColor Yellow
                Add-Result 'OneDrive' (Split-Path $Target.Path -Leaf) $Status
                continue
            }

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
                39 { return 'RIGHT' }
                37 { return 'LEFT' }
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

function Invoke-Reboot {
    Write-Host ''
    Write-Host "  $($T.Rebooting)" -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    Restart-Computer -Force
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

function Get-Component {
    param([int]$Index)

    switch ($Index) {
        1 { return @{ Key='Office';   Label=$T.OptOffice;   Items=$Script:Office;   Picks=$Script:OfficePick } }
        2 { return @{ Key='Copilot';  Label=$T.OptCopilot;  Items=$Script:Copilot;  Picks=$Script:CopilotPick } }
        3 { return @{ Key='OneDrive'; Label=$T.OptOneDrive; Items=$Script:OneDrive; Picks=$Script:OneDrivePick } }
        4 { return @{ Key='Hp';       Label=$T.OptHp;       Items=$Script:Hp;       Picks=$Script:HpPick } }
    }
    return $null
}

function Get-ItemLabel {
    param([int]$Comp, $Item)

    switch ($Comp) {
        2 { return "$($Item.Name)  ($($Item.Kind))" }
        3 { return $Item.Path }
        default { return $Item.DisplayName }
    }
}

function Get-ItemTag {
    param($Item)

    if ($Item.PSObject.Properties['Keep']     -and $Item.Keep)     { return $T.TagKeep }
    if ($Item.PSObject.Properties['Licensed'] -and $Item.Licensed) { return $T.TagLicensed }
    return ''
}

function Get-TickedCount {
    param([array]$Items, [array]$Picks)

    $n = 0
    for ($i = 0; $i -lt $Items.Count; $i++) { if ($Picks[$i]) { $n++ } }
    return $n
}

function Get-CompState {
    # none / partial / all - drives the checkbox glyph
    param([int]$Comp)

    $Cfg = Get-Component -Index $Comp
    if ($Cfg.Items.Count -eq 0) {
        return $(if ($Sel[$Cfg.Key]) { 'all' } else { 'none' })
    }

    $Ticked = Get-TickedCount -Items $Cfg.Items -Picks $Cfg.Picks
    if ($Ticked -eq 0) { return 'none' }

    # An item flagged Keep is not part of a "full" selection, so a group whose
    # only unticked entry is HP Support Assistant still counts as fully on.
    for ($i = 0; $i -lt $Cfg.Items.Count; $i++) {
        if ((-not $Cfg.Items[$i].Keep) -and (-not $Cfg.Picks[$i])) { return 'partial' }
    }
    return 'all'
}

function Get-CompChecked {
    param([int]$Comp)
    return ((Get-CompState -Comp $Comp) -ne 'none')
}

function Switch-Comp {
    param([int]$Comp, [string]$Force = '')

    $Cfg = Get-Component -Index $Comp

    # A component with nothing found is a plain on/off flag: Copilot still
    # applies its policy, so it stays tickable.
    if ($Cfg.Items.Count -eq 0) {
        $Sel[$Cfg.Key] = switch ($Force) { 'on' { $true } 'off' { $false } default { -not $Sel[$Cfg.Key] } }
        return
    }

    $TurnOn = switch ($Force) {
        'on'    { $true }
        'off'   { $false }
        default { (Get-CompState -Comp $Comp) -eq 'none' }
    }

    for ($i = 0; $i -lt $Cfg.Items.Count; $i++) {
        # Ticking the group never sweeps up an item flagged Keep
        $Cfg.Picks[$i] = $TurnOn -and (-not $Cfg.Items[$i].Keep)
    }
    $Sel[$Cfg.Key] = $TurnOn
}

function Switch-Item {
    param([int]$Comp, [int]$Index)

    $Cfg = Get-Component -Index $Comp
    $Cfg.Picks[$Index] = -not $Cfg.Picks[$Index]
    $Sel[$Cfg.Key]     = ((Get-TickedCount -Items $Cfg.Items -Picks $Cfg.Picks) -gt 0)
}

function Get-AllState {
    # A component with nothing found cannot be "on" in any meaningful sense,
    # so it does not drag the group header into a partial state.
    $States = @(1..4 |
        Where-Object { (Get-Component -Index $_).Items.Count -gt 0 } |
        ForEach-Object { Get-CompState -Comp $_ })

    if ($States.Count -eq 0)                                      { return 'none' }
    if ($States -notcontains 'all' -and $States -notcontains 'partial') { return 'none' }
    if ($States -contains 'none' -or $States -contains 'partial') { return 'partial' }
    return 'all'
}

function Switch-All {
    $Target = if ((Get-AllState) -eq 'none') { 'on' } else { 'off' }
    foreach ($c in 1..4) { Switch-Comp -Comp $c -Force $Target }
}

function Get-VisibleRows {
    # The tree, flattened: what the cursor actually walks over
    $Rows = @()
    $Rows += @{ Type = 'all' }

    foreach ($c in 1..4) {
        $Rows += @{ Type = 'comp'; Comp = $c }

        if ($Script:Expanded[$c]) {
            $Cfg = Get-Component -Index $c
            for ($i = 0; $i -lt $Cfg.Items.Count; $i++) {
                $Rows += @{ Type = 'item'; Comp = $c; Index = $i }
            }
        }
    }
    return $Rows
}

function Get-Box {
    param([string]$State)

    switch ($State) {
        'all'     { return '[x]' }
        'partial' { return '[~]' }
        default   { return '[ ]' }
    }
}

function Show-Tree {
    param([array]$Rows, [int]$Cursor)

    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $Row    = $Rows[$r]
        $IsHere = ($r -eq $Cursor)
        $Arrow  = if ($IsHere) { '>' } else { ' ' }

        if ($Row.Type -eq 'all') {
            $State = Get-AllState
            $Col   = if ($IsHere) { 'Cyan' } elseif ($State -eq 'none') { 'Gray' } else { 'Green' }
            Write-Host ('   {0}  {1}  {2}' -f $Arrow, (Get-Box $State), $T.OptAll) -ForegroundColor $Col
            Write-Host '      ----------------------------------------------------------' -ForegroundColor DarkGray
            continue
        }

        if ($Row.Type -eq 'comp') {
            $Cfg   = Get-Component -Index $Row.Comp
            $State = Get-CompState -Comp $Row.Comp
            $Col   = if ($IsHere) { 'Cyan' } elseif ($State -eq 'none') { 'Gray' } else { 'Green' }

            $Chev = '   '
            if ($Cfg.Items.Count -gt 0) {
                $Chev = if ($Script:Expanded[$Row.Comp]) { ' v ' } else { ' > ' }
            }

            Write-Host ('   {0}  {1} {2} {3,-24} {4}' -f `
                $Arrow, (Get-Box $State), $Chev, $Cfg.Label, `
                (Get-ComponentDetail -Comp $Row.Comp)) -ForegroundColor $Col
            continue
        }

        # item
        $Cfg  = Get-Component -Index $Row.Comp
        $Item = $Cfg.Items[$Row.Index]
        $On   = [bool]$Cfg.Picks[$Row.Index]
        $Tag  = Get-ItemTag -Item $Item

        $Col = if ($IsHere) { 'Cyan' } elseif ($On) { 'Green' } else { 'DarkGray' }
        $Txt = '{0}        {1}  {2}' -f $Arrow, $(if ($On) { '[x]' } else { '[ ]' }), (Get-ItemLabel -Comp $Row.Comp -Item $Item)
        if ($Tag) { $Txt = "$Txt   $Tag" }

        Write-Host ("   $Txt") -ForegroundColor $Col
    }
}

function Get-ComponentDetail {
    param([int]$Comp)

    $Cfg = Get-Component -Index $Comp
    if ($Cfg.Items.Count -eq 0) { return $T.NotFound }

    $Ticked = Get-TickedCount -Items $Cfg.Items -Picks $Cfg.Picks
    $Text   = $T.FoundCount -f $Cfg.Items.Count
    if ($Ticked -ne $Cfg.Items.Count) {
        $Text = "$Text  " + ($T.SelCount -f $Ticked)
    }
    return $Text
}

function Get-Picked {
    param([array]$Items, [array]$Picks)

    $Picked = @()
    for ($i = 0; $i -lt $Items.Count; $i++) {
        if ($Picks[$i]) { $Picked += $Items[$i] }
    }
    $Picked
}

function Select-Language {
    Show-Banner
    Write-Host $T.LangHeader -ForegroundColor White
    Write-Host ''
    Write-Host '   [1] English'    -ForegroundColor Gray
    Write-Host '   [2] Nederlands' -ForegroundColor Gray
    Write-Host '   [3] Francais'   -ForegroundColor Gray
    Write-Host ''

    switch (Read-MenuKey) {
        '1' { return 'EN' }
        '2' { return 'NL' }
        '3' { return 'FR' }
        default { return $script:Lang }
    }
}


# ------------------------------------------------------------------- Main ---

Disable-QuickEdit
Show-Banner
Invoke-Scan

$Cursor = 0

do {
    $Rows = @(Get-VisibleRows)
    if ($Cursor -ge $Rows.Count) { $Cursor = $Rows.Count - 1 }
    if ($Cursor -lt 0)           { $Cursor = 0 }

    Show-Banner
    Write-Host "  $($T.SelectHeader)" -ForegroundColor White
    Write-Host ''

    Show-Tree -Rows $Rows -Cursor $Cursor
    Write-Host ''

    if ($Script:RawKeys) {
        Write-Host "   $($T.NavHint)" -ForegroundColor DarkGray
    }
    else {
        Write-Host "   $($T.NavHintKeys)" -ForegroundColor DarkGray
    }
    Write-Host ''

    Write-Host "  $($T.MenuStart)"  -ForegroundColor White
    Write-Host "  $($T.MenuRescan)" -ForegroundColor White
    Write-Host "  $($T.MenuLang) - $($T.LangLabel)" -ForegroundColor White
    Write-Host "  $($T.MenuReboot)" -ForegroundColor White
    Write-Host "  $($T.MenuExit)"   -ForegroundColor White
    Write-Host ''

    $Choice = Read-MenuKey
    $Row    = $Rows[$Cursor]

    switch ($Choice) {

        'UP'   { $Cursor--; if ($Cursor -lt 0)           { $Cursor = $Rows.Count - 1 } }
        'DOWN' { $Cursor++; if ($Cursor -ge $Rows.Count) { $Cursor = 0 } }

        'SPACE' {
            switch ($Row.Type) {
                'all'  { Switch-All }
                'comp' { Switch-Comp -Comp $Row.Comp }
                'item' { Switch-Item -Comp $Row.Comp -Index $Row.Index }
            }
        }

        { $_ -in 'RIGHT', 'ENTER', 'V' } {
            if ($Row.Type -eq 'comp') {
                $Cfg = Get-Component -Index $Row.Comp
                if ($Cfg.Items.Count -gt 0) {
                    if ($Script:Expanded[$Row.Comp]) {
                        $Cursor++        # already open: step into the group
                    }
                    else {
                        $Script:Expanded[$Row.Comp] = $true
                    }
                }
            }
        }

        'LEFT' {
            if ($Row.Type -ne 'all') {
                $Target = $Row.Comp
                if ($Row.Type -eq 'item' -or $Script:Expanded[$Target]) {
                    $Script:Expanded[$Target] = $false

                    # park the cursor back on the group header
                    $Back = @(Get-VisibleRows)
                    for ($i = 0; $i -lt $Back.Count; $i++) {
                        if ($Back[$i].Type -eq 'comp' -and $Back[$i].Comp -eq $Target) { $Cursor = $i; break }
                    }
                }
            }
        }

        # Letter shortcuts, mainly for the typed fallback
        'A' { Switch-All }
        'O' { Switch-Comp -Comp 1 }
        'C' { Switch-Comp -Comp 2 }
        'D' { Switch-Comp -Comp 3 }
        'H' { Switch-Comp -Comp 4 }

        '2' {
            Show-Banner
            Invoke-Scan
        }

        '3' {
            $Lang = Select-Language
            $T    = $Strings[$Lang]
            $Host.UI.RawUI.WindowTitle = $T.WindowTitle
        }

        '4' {
            Write-Host ''
            $Answer = Read-Host $T.ConfirmReboot
            if ($Answer -match $T.YesPattern) { Invoke-Reboot }
        }

        '1' {

            $DoOffice   = Get-CompChecked -Comp 1
            $DoCopilot  = Get-CompChecked -Comp 2
            $DoOneDrive = Get-CompChecked -Comp 3
            $DoHp       = Get-CompChecked -Comp 4

            if (-not ($DoOffice -or $DoCopilot -or $DoOneDrive -or $DoHp)) {
                Write-Host ''
                Write-Host "  $($T.NothingSel)" -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                continue
            }

            $OfficeToRemove   = @(Get-Picked -Items $Office   -Picks $Script:OfficePick)
            $CopilotToRemove  = @(Get-Picked -Items $Copilot  -Picks $Script:CopilotPick)
            $OneDriveToRemove = @(Get-Picked -Items $OneDrive -Picks $Script:OneDrivePick)
            $HpToRemove       = @(Get-Picked -Items $Hp       -Picks $Script:HpPick)

            # A selected component with nothing ticked simply contributes
            # nothing; only a zero total is worth stopping for.
            $Planned = 0
            if ($DoOffice)   { $Planned += $OfficeToRemove.Count }
            if ($DoCopilot)  { $Planned += $CopilotToRemove.Count }
            if ($DoOneDrive) { $Planned += $OneDriveToRemove.Count }
            if ($DoHp)       { $Planned += $HpToRemove.Count }

            if ($Planned -eq 0) {
                $AnyFound = ($DoOffice   -and $Office.Count   -gt 0) -or
                            ($DoCopilot  -and $Copilot.Count  -gt 0) -or
                            ($DoOneDrive -and $OneDrive.Count -gt 0) -or
                            ($DoHp       -and $Hp.Count       -gt 0)

                Write-Host ''
                if ($AnyFound) {
                    # Things were detected, they are just all unticked
                    Write-Host "  $($T.NoneTicked)" -ForegroundColor Yellow
                }
                else {
                    Write-Host "  $($T.NothingFound)" -ForegroundColor Yellow
                }
                Start-Sleep -Seconds 3
                continue
            }

            Write-Host ''
            if ($DoOneDrive -and $OneDriveToRemove.Count -gt 0) {
                Write-Host "  $($T.WarnOneDrive)" -ForegroundColor Yellow
                Write-Host ''
            }
            if ($DoHp -and ($HpToRemove | Where-Object { $_.DisplayName -match 'Wolf Security|Sure Click|Sure Sense|Client Security' })) {
                Write-Host "  $($T.WarnHp)" -ForegroundColor Yellow
                Write-Host ''
            }

            $LicensedTicked = @($OfficeToRemove | Where-Object { $_.Licensed })
            if ($DoOffice -and $LicensedTicked.Count -gt 0) {
                Write-Host "  $($T.WarnLicensed)" -ForegroundColor Red
                foreach ($l in $LicensedTicked) {
                    Write-Host "    - $($l.DisplayName)" -ForegroundColor Red
                }
                Write-Host "  $($T.WarnLicensed2)" -ForegroundColor Yellow
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
            $Script:HpStuck      = $false

            Write-Host ''
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host $T.HeaderRemoval -ForegroundColor White
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host ''

            if ($DoOffice)   { Remove-OfficeTargets   -Targets $OfficeToRemove }
            if ($DoCopilot)  { Remove-CopilotTargets  -Targets $CopilotToRemove }
            if ($DoOneDrive) { Remove-OneDriveTargets -Targets $OneDriveToRemove }
            if ($DoHp)       { Remove-HpTargets       -Targets $HpToRemove }

            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            Write-Host $T.HeaderSummary -ForegroundColor White
            Write-Host '  ------------------------------------------------------------' -ForegroundColor DarkGray
            foreach ($r in $Script:Results) {
                Write-Host ('    {0,-9} {1,-32} {2}' -f $r.Component, $r.Name, $r.Status) -ForegroundColor Gray
            }
            Write-Host ''

            if ($Script:HpStuck) {
                Write-Host "  $($T.HpStuckNote)" -ForegroundColor Yellow
                Write-Host ''
            }

            if ($Script:RebootNeeded) {
                Write-Host "  $($T.RebootNote)" -ForegroundColor Yellow
            }
            else {
                Write-Host "  $($T.RunDone)" -ForegroundColor Green
            }
            Write-Host ''

            # A reboot is always on offer here: HP uninstallers leave the shell
            # and various services in a half-restarted state.
            $Answer = Read-Host $T.RebootAsk
            if ($Answer -match $T.YesPattern) { Invoke-Reboot }
            Write-Host ''

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
