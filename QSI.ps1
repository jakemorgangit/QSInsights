#Requires -Version 5.1
<#
.SYNOPSIS
QS Insights

.DESCRIPTION
GUI tool to explore SQL Server query plans using Query Store.

Phase 1:
- Load longest running queries from Query Store without doing XML analysis on the server.

Phase 2:
- Analyse the locally loaded execution plans for issues (spills, missing indexes, memory grants, etc)
  when the user clicks the green Insights button.

Additional features:
- Session profiles saved in Documents\QSInsights\Sessions.xml
- Dark themed WinForms UI
- Time window presets (1 hour to 4 weeks)
- Database selector listing only databases with Query Store enabled
- Options to exclude index creation and stats updates
- Top N selector
- Grid of results with:
  - last_duration (microseconds), seconds and HH:MM:SS.mmm
  - query_hash_hex and (if available) query_plan_hash_hex
  - query text, plan, last_execution_time
  - Plan Insights (missing index, implicit conversion, spills, memory grants, row goals,
    no join predicate, plan affecting converts, missing stats, CE warnings, non parallel reasons, other warnings)
  - Copy row to clipboard
  - Open execution plan (.sqlplan)
  - Row highlighting based on Plan Insights
- Save results to JSON and load from JSON (no DB hit)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

# Ensure STA
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    [System.Windows.Forms.MessageBox]::Show(
        "Restarting QS Insights in STA mode for the GUI...",
        "QS Insights",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null

    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -WindowStyle Hidden
    exit
}

# ====================================================================================
# SESSION STORAGE
# ====================================================================================

$script:RootFolder       = Join-Path $env:USERPROFILE "Documents\QSInsights"
$script:SessionsFilePath = Join-Path $script:RootFolder "Sessions.xml"
$script:Sessions         = New-Object System.Collections.ArrayList
$script:CurrentSession   = $null

function Ensure-SessionsStorage {
    if (-not (Test-Path $script:RootFolder)) {
        New-Item -ItemType Directory -Path $script:RootFolder -Force | Out-Null
    }
    if (-not (Test-Path $script:SessionsFilePath)) {
        $blank = [pscustomobject]@{
            LastSession = ""
            Sessions    = @()
        }
        $blank | Export-Clixml -Path $script:SessionsFilePath
    }
}

function Convert-ToArrayList {
    param($items)
    $list = New-Object System.Collections.ArrayList
    if ($null -eq $items) { return $list }

    if ($items -is [System.Collections.IEnumerable] -and
        -not ($items -is [string])) {
        foreach ($i in $items) { [void]$list.Add($i) }
    }
    else {
        [void]$list.Add($items)
    }
    return $list
}

function Load-AllSessions {
    Ensure-SessionsStorage
    try {
        $data = Import-Clixml -Path $script:SessionsFilePath
        $script:Sessions = New-Object System.Collections.ArrayList

        if ($data.Sessions) {
            $tmp = Convert-ToArrayList $data.Sessions
            foreach ($s in $tmp) {
                $null = $script:Sessions.Add(
                    [pscustomobject]@{
                        Name     = $s.Name
                        Server   = $s.Server
                        Username = $s.Username
                        Password = $s.Password
                    }
                )
            }
        }
        $script:CurrentSession = $data.LastSession
    }
    catch {
        $script:Sessions       = New-Object System.Collections.ArrayList
        $script:CurrentSession = $null
    }
}

function Save-AllSessions {
    Ensure-SessionsStorage
    $plainSessions = @()
    foreach ($s in $script:Sessions) {
        $plainSessions += [pscustomobject]@{
            Name     = $s.Name
            Server   = $s.Server
            Username = $s.Username
            Password = $s.Password
        }
    }
    [pscustomobject]@{
        LastSession = $script:CurrentSession
        Sessions    = $plainSessions
    } | Export-Clixml -Path $script:SessionsFilePath
}

function Ensure-SessionsArrayList {
    if (-not ($script:Sessions -is [System.Collections.ArrayList])) {
        $script:Sessions = Convert-ToArrayList $script:Sessions
    }
}

function Get-SessionByName {
    param([string]$Name)
    foreach ($s in $script:Sessions) {
        if ($s.Name -eq $Name) { return $s }
    }
    return $null
}

function Upsert-Session {
    param(
        [string]$Name,
        [string]$Server,
        [string]$Username,
        [securestring]$Password
    )

    Ensure-SessionsArrayList

    $existing = Get-SessionByName -Name $Name
    if ($existing) {
        $existing.Server   = $Server
        $existing.Username = $Username
        $existing.Password = $Password
    } else {
        $null = $script:Sessions.Add(
            [pscustomobject]@{
                Name     = $Name
                Server   = $Server
                Username = $Username
                Password = $Password
            }
        )
    }

    $script:CurrentSession = $Name
    Save-AllSessions
}

function Remove-SessionByName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return }

    Ensure-SessionsArrayList

    $newList = New-Object System.Collections.ArrayList
    foreach ($sess in $script:Sessions) {
        if ($sess.Name -ne $Name) {
            $null = $newList.Add($sess)
        }
    }
    $script:Sessions = $newList

    if ($script:CurrentSession -eq $Name) {
        $script:CurrentSession = $null
    }

    Save-AllSessions
}

# ====================================================================================
# THEME
# ====================================================================================

$bgMain        = [System.Drawing.ColorTranslator]::FromHtml("#1e1e1e")
$bgPanel       = [System.Drawing.ColorTranslator]::FromHtml("#252526")
$bgPanelBorder = [System.Drawing.ColorTranslator]::FromHtml("#3f3f46")

$fgPrimary     = [System.Drawing.ColorTranslator]::FromHtml("#d4d4d4")
$fgSecondary   = [System.Drawing.ColorTranslator]::FromHtml("#9ca3af")
$fgAccent      = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")

$accentBlue    = [System.Drawing.ColorTranslator]::FromHtml("#0e639c")
$btnGray       = [System.Drawing.ColorTranslator]::FromHtml("#3a3d41")
$blockingRed   = [System.Drawing.ColorTranslator]::FromHtml("#b00020")
$accentGreen   = [System.Drawing.ColorTranslator]::FromHtml("#22c55e")

$gridHeaderBg  = [System.Drawing.ColorTranslator]::FromHtml("#2d2d30")
$gridHeaderFg  = $fgAccent
$gridRowBg     = [System.Drawing.ColorTranslator]::FromHtml("#1e1e1e")
$gridRowSelBg  = [System.Drawing.ColorTranslator]::FromHtml("#094771")
$gridRowSelFg  = $fgAccent
$gridLines     = [System.Drawing.ColorTranslator]::FromHtml("#3f3f46")

$consoleBack   = [System.Drawing.ColorTranslator]::FromHtml("#1a1a1a")
$consoleFore   = [System.Drawing.ColorTranslator]::FromHtml("#d4d4d4")

$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$fontBold    = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

function New-DarkGroupBox {
    param(
        [string]$text,
        [System.Drawing.Point]$location,
        [System.Drawing.Size]$size,
        [string]$anchor = "Top, Left"
    )

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text      = $text
    $gb.Location  = $location
    $gb.Size      = $size
    $gb.Anchor    = $anchor
    $gb.ForeColor = $fgPrimary
    $gb.BackColor = $bgPanel
    return $gb
}

function Apply-DarkModeToGrid {
    param([System.Windows.Forms.DataGridView]$grid)

    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersBorderStyle  = 'None'
    $grid.RowHeadersVisible         = $false

    $headerStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $headerStyle.BackColor = $gridHeaderBg
    $headerStyle.ForeColor = $gridHeaderFg
    $headerStyle.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle = $headerStyle

    $cellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $cellStyle.BackColor          = $gridRowBg
    $cellStyle.ForeColor          = $fgPrimary
    $cellStyle.SelectionBackColor = $gridRowSelBg
    $cellStyle.SelectionForeColor = $gridRowSelFg

    $grid.DefaultCellStyle = $cellStyle
    $grid.BackgroundColor  = $gridRowBg
    $grid.GridColor        = $gridLines
    $grid.BorderStyle      = 'None'
    $grid.ReadOnly         = $true
    $grid.SelectionMode    = 'FullRowSelect'
    $grid.MultiSelect      = $false
    $grid.Font             = $fontRegular
}

# ====================================================================================
# HELPERS
# ====================================================================================

function Log-Message {
    param([string]$Message)
    $consoleLogTextBox.AppendText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n")
    $consoleLogTextBox.Update()
}

function Get-PasswordFromTextBox {
    if ($passwordTextBox.Tag -is [System.Security.SecureString]) {
        $bstr  = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordTextBox.Tag)
        $plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        return $plain
    } else {
        return $passwordTextBox.Text
    }
}

function Get-SqlData {
    param(
        [string]$ConnectionString,
        [string]$Query
    )

    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $command    = New-Object System.Data.SqlClient.SqlCommand($Query, $connection)
    $command.CommandTimeout = 300

    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
    $dataset = New-Object System.Data.DataSet

    try {
        $connection.Open()
        [void]$adapter.Fill($dataset)
    }
    catch {
        throw $_
    }
    finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
    }

    $table = $null
    if ($dataset.Tables.Count -gt 0) {
        $table = ($dataset.Tables | Where-Object { $_.Columns.Count -gt 0 } | Select-Object -Last 1)
    }
    if (-not $table) { return @() }

    $results = New-Object System.Collections.ArrayList
    foreach ($row in $table.Rows) {
        $obj = New-Object -TypeName PSObject
        foreach ($col in $table.Columns) {
            $obj | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value $row[$col]
        }
        [void]$results.Add($obj)
    }
    return $results
}

function Get-QueryStoreTimeWindowExpression {
    $choice = [string]$timeWindowCombo.SelectedItem
    switch ($choice) {
        'Last 1 hour'   { 'DATEADD(HOUR, -1, SYSUTCDATETIME())' }
        'Last 2 hours'  { 'DATEADD(HOUR, -2, SYSUTCDATETIME())' }
        'Last 6 hours'  { 'DATEADD(HOUR, -6, SYSUTCDATETIME())' }
        'Last 12 hours' { 'DATEADD(HOUR, -12, SYSUTCDATETIME())' }
        'Last 1 day'    { 'DATEADD(DAY,  -1, SYSUTCDATETIME())' }
        'Last 2 days'   { 'DATEADD(DAY,  -2, SYSUTCDATETIME())' }
        'Last 5 days'   { 'DATEADD(DAY,  -5, SYSUTCDATETIME())' }
        'Last 7 days'   { 'DATEADD(DAY,  -7, SYSUTCDATETIME())' }
        'Last 10 days'  { 'DATEADD(DAY, -10, SYSUTCDATETIME())' }
        'Last 2 weeks'  { 'DATEADD(DAY, -14, SYSUTCDATETIME())' }
        'Last 4 weeks'  { 'DATEADD(DAY, -28, SYSUTCDATETIME())' }
        default         { 'DATEADD(HOUR, -1, SYSUTCDATETIME())' }
    }
}

# Load DBs with Query Store enabled
function Load-QueryStoreDatabases {
    param(
        [string]$Server,
        [string]$Username,
        [string]$Password
    )

    $databaseComboBox.Items.Clear()
    $databaseComboBox.Enabled = $false

    Log-Message "Loading databases with Query Store enabled..."

    $csMaster = "Server=$Server;Database=master;User ID=$Username;Password=$Password;TrustServerCertificate=True"

    $qsDbs = New-Object System.Collections.ArrayList
    $usedSysDatabases = $false

    try {
        $rows = Get-SqlData -ConnectionString $csMaster -Query @"
SELECT name
FROM sys.databases
WHERE state_desc = 'ONLINE'
  AND database_id > 4
  AND is_query_store_on = 1
ORDER BY name;
"@
        $usedSysDatabases = $true

        if ($rows -and $rows.Count -gt 0) {
            foreach ($r in $rows) {
                [void]$qsDbs.Add([string]$r.name)
            }
            Log-Message ("Found {0} database(s) with Query Store enabled (via sys.databases)." -f $qsDbs.Count)
        }
        else {
            Log-Message "sys.databases reports no Query Store databases; falling back to per database check..."
        }
    }
    catch {
        Log-Message "sys.databases check for Query Store failed: $($_.Exception.Message). Falling back to per database check..."
    }

    if ($qsDbs.Count -eq 0) {
        try {
            $dbRows = Get-SqlData -ConnectionString $csMaster -Query @"
SELECT name
FROM sys.databases
WHERE state_desc = 'ONLINE'
  AND database_id > 4
ORDER BY name;
"@

            foreach ($db in $dbRows) {
                $dbName = [string]$db.name
                $csDb   = "Server=$Server;Database=$dbName;User ID=$Username;Password=$Password;TrustServerCertificate=True"
                try {
                    $rows = Get-SqlData -ConnectionString $csDb -Query "SELECT actual_state_desc FROM sys.database_query_store_options;"
                    if ($rows -and $rows.Count -gt 0) {
                        $state = [string]$rows[0].actual_state_desc
                        if ($state -ne 'OFF') {
                            [void]$qsDbs.Add($dbName)
                        }
                    }
                }
                catch {
                    Log-Message "Warning: could not read Query Store options for database '$dbName' ($($_.Exception.Message))"
                }
            }

            if ($qsDbs.Count -gt 0) {
                if ($usedSysDatabases) {
                    Log-Message ("Per database fallback found {0} Query Store database(s)." -f $qsDbs.Count)
                }
                else {
                    Log-Message ("Found {0} database(s) with Query Store enabled (per database check)." -f $qsDbs.Count)
                }
            }
            else {
                Log-Message "No databases with Query Store enabled were found."
            }
        }
        catch {
            Log-Message "ERROR: Failed to enumerate databases for Query Store check. $($_.Exception.Message)"
        }
    }

    if ($qsDbs.Count -gt 0) {
        $databaseComboBox.BeginUpdate()
        foreach ($n in ($qsDbs | Sort-Object)) {
            [void]$databaseComboBox.Items.Add($n)
        }
        $databaseComboBox.EndUpdate()
        $databaseComboBox.Enabled       = $true
        $databaseComboBox.SelectedIndex = 0
    }
}

# ====================================================================================
# MAIN FORM
# ====================================================================================

[System.Windows.Forms.Application]::EnableVisualStyles()

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text          = "QS Insights"
$mainForm.Size          = New-Object System.Drawing.Size(1150, 750)
$mainForm.MinimumSize   = New-Object System.Drawing.Size(900, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.BackColor     = $bgMain
$mainForm.ForeColor     = $fgPrimary
$mainForm.Font          = $fontRegular

# Menu
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.BackColor = $bgPanel
$menuStrip.ForeColor = $fgPrimary

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("File")
$miExit   = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
[void]$fileMenu.DropDownItems.Add($miExit)

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
$miAbout  = New-Object System.Windows.Forms.ToolStripMenuItem("About")
[void]$helpMenu.DropDownItems.Add($miAbout)

[void]$menuStrip.Items.AddRange(@($fileMenu, $helpMenu))
$mainForm.MainMenuStrip = $menuStrip
$mainForm.Controls.Add($menuStrip)

# ====================================================================================
# CONNECTION DETAILS GROUP
# ====================================================================================

$connectionGroupBox = New-DarkGroupBox -text "Connection Details" `
    -location ([System.Drawing.Point]::new(20, 40)) `
    -size     ([System.Drawing.Size]::new(300, 290)) `
    -anchor   "Top, Left"

$mainForm.Controls.Add($connectionGroupBox)

$sessionLabel = New-Object System.Windows.Forms.Label
$sessionLabel.Text      = "Session:"
$sessionLabel.Location  = New-Object System.Drawing.Point(10, 30)
$sessionLabel.AutoSize  = $true
$sessionLabel.ForeColor = $fgPrimary

$sessionComboBox = New-Object System.Windows.Forms.ComboBox
$sessionComboBox.Location      = New-Object System.Drawing.Point(120, 27)
$sessionComboBox.Size          = New-Object System.Drawing.Size(160, 20)
$sessionComboBox.DropDownStyle = 'DropDownList'
$sessionComboBox.BackColor     = $gridRowBg
$sessionComboBox.ForeColor     = $fgPrimary
$sessionComboBox.FlatStyle     = 'Flat'

$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Text      = "Server/Instance:"
$serverLabel.Location  = New-Object System.Drawing.Point(10, 60)
$serverLabel.AutoSize  = $true
$serverLabel.ForeColor = $fgPrimary

$serverTextBox = New-Object System.Windows.Forms.TextBox
$serverTextBox.Location    = New-Object System.Drawing.Point(120, 57)
$serverTextBox.Size        = New-Object System.Drawing.Size(160, 20)
$serverTextBox.BackColor   = $gridRowBg
$serverTextBox.ForeColor   = $fgPrimary
$serverTextBox.BorderStyle = 'FixedSingle'

$usernameLabel = New-Object System.Windows.Forms.Label
$usernameLabel.Text      = "Username:"
$usernameLabel.Location  = New-Object System.Drawing.Point(10, 90)
$usernameLabel.AutoSize  = $true
$usernameLabel.ForeColor = $fgPrimary

$usernameTextBox = New-Object System.Windows.Forms.TextBox
$usernameTextBox.Location    = New-Object System.Drawing.Point(120, 87)
$usernameTextBox.Size        = New-Object System.Drawing.Size(160, 20)
$usernameTextBox.BackColor   = $gridRowBg
$usernameTextBox.ForeColor   = $fgPrimary
$usernameTextBox.BorderStyle = 'FixedSingle'

$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Text      = "Password:"
$passwordLabel.Location  = New-Object System.Drawing.Point(10, 120)
$passwordLabel.AutoSize  = $true
$passwordLabel.ForeColor = $fgPrimary

$passwordTextBox = New-Object System.Windows.Forms.TextBox
$passwordTextBox.Location   = New-Object System.Drawing.Point(120, 117)
$passwordTextBox.Size       = New-Object System.Drawing.Size(160, 20)
$passwordTextBox.UseSystemPasswordChar = $true
$passwordTextBox.BackColor  = $gridRowBg
$passwordTextBox.ForeColor  = $fgPrimary
$passwordTextBox.BorderStyle = 'FixedSingle'

$testConnectionButton = New-Object System.Windows.Forms.Button
$testConnectionButton.Text      = "Test"
$testConnectionButton.Location  = New-Object System.Drawing.Point(10, 155)
$testConnectionButton.Size      = New-Object System.Drawing.Size(80, 30)
$testConnectionButton.BackColor = $btnGray
$testConnectionButton.ForeColor = $fgPrimary
$testConnectionButton.FlatStyle = 'Flat'

$saveConnectionButton = New-Object System.Windows.Forms.Button
$saveConnectionButton.Text      = "Save"
$saveConnectionButton.Location  = New-Object System.Drawing.Point(100, 155)
$saveConnectionButton.Size      = New-Object System.Drawing.Size(80, 30)
$saveConnectionButton.BackColor = $btnGray
$saveConnectionButton.ForeColor = $fgPrimary
$saveConnectionButton.FlatStyle = 'Flat'

$deleteSessionButton = New-Object System.Windows.Forms.Button
$deleteSessionButton.Text      = "🗑 Delete"
$deleteSessionButton.Location  = New-Object System.Drawing.Point(190, 155)
$deleteSessionButton.Size      = New-Object System.Drawing.Size(80, 30)
$deleteSessionButton.BackColor = $btnGray
$deleteSessionButton.ForeColor = $blockingRed
$deleteSessionButton.FlatStyle = 'Flat'
$deleteSessionButton.Font      = New-Object System.Drawing.Font("Segoe UI Emoji", 9, [System.Drawing.FontStyle]::Regular)

$connectionStatusLabel = New-Object System.Windows.Forms.Label
$connectionStatusLabel.Text      = "Status: Not Connected"
$connectionStatusLabel.Location  = New-Object System.Drawing.Point(10, 195)
$connectionStatusLabel.AutoSize  = $true
$connectionStatusLabel.ForeColor = $fgSecondary

$databaseLabel = New-Object System.Windows.Forms.Label
$databaseLabel.Text      = "Database (Query Store):"
$databaseLabel.Location  = New-Object System.Drawing.Point(10, 220)
$databaseLabel.AutoSize  = $true
$databaseLabel.ForeColor = $fgPrimary

$databaseComboBox = New-Object System.Windows.Forms.ComboBox
$databaseComboBox.Location      = New-Object System.Drawing.Point(150, 217)
$databaseComboBox.Size          = New-Object System.Drawing.Size(130, 20)
$databaseComboBox.DropDownStyle = 'DropDownList'
$databaseComboBox.BackColor     = $gridRowBg
$databaseComboBox.ForeColor     = $fgPrimary
$databaseComboBox.FlatStyle     = 'Flat'
$databaseComboBox.Enabled       = $false

$currentSessionLabel = New-Object System.Windows.Forms.Label
$currentSessionLabel.Text      = "Current Session:"
$currentSessionLabel.Location  = New-Object System.Drawing.Point(10, 250)
$currentSessionLabel.AutoSize  = $true
$currentSessionLabel.ForeColor = $fgPrimary

$currentSessionValue = New-Object System.Windows.Forms.Label
$currentSessionValue.Text      = "(none)"
$currentSessionValue.Location  = New-Object System.Drawing.Point(120, 250)
$currentSessionValue.AutoSize  = $true
$currentSessionValue.ForeColor = $fgPrimary

$connectionGroupBox.Controls.AddRange(@(
    $sessionLabel, $sessionComboBox,
    $serverLabel, $serverTextBox,
    $usernameLabel, $usernameTextBox,
    $passwordLabel, $passwordTextBox,
    $testConnectionButton, $saveConnectionButton, $deleteSessionButton,
    $connectionStatusLabel,
    $databaseLabel, $databaseComboBox,
    $currentSessionLabel, $currentSessionValue
))

# ====================================================================================
# QUERY FILTERS GROUP
# ====================================================================================

$queryGroupBox = New-DarkGroupBox -text "Query Filters" `
    -location ([System.Drawing.Point]::new(20, 340)) `
    -size     ([System.Drawing.Size]::new(300, 270)) `
    -anchor   "Top, Left"

$mainForm.Controls.Add($queryGroupBox)

$timeWindowLabel = New-Object System.Windows.Forms.Label
$timeWindowLabel.Text      = "Time window:"
$timeWindowLabel.Location  = New-Object System.Drawing.Point(10, 35)
$timeWindowLabel.AutoSize  = $true
$timeWindowLabel.ForeColor = $fgPrimary

$timeWindowCombo = New-Object System.Windows.Forms.ComboBox
$timeWindowCombo.Location      = New-Object System.Drawing.Point(120, 32)
$timeWindowCombo.Size          = New-Object System.Drawing.Size(160, 20)
$timeWindowCombo.DropDownStyle = 'DropDownList'
$timeWindowCombo.BackColor     = $gridRowBg
$timeWindowCombo.ForeColor     = $fgPrimary
$timeWindowCombo.FlatStyle     = 'Flat'

$timeWindowCombo.Items.AddRange(@(
    "Last 1 hour",
    "Last 2 hours",
    "Last 6 hours",
    "Last 12 hours",
    "Last 1 day",
    "Last 2 days",
    "Last 5 days",
    "Last 7 days",
    "Last 10 days",
    "Last 2 weeks",
    "Last 4 weeks"
)) | Out-Null
$timeWindowCombo.SelectedIndex = 7

$excludeIndexCheckBox = New-Object System.Windows.Forms.CheckBox
$excludeIndexCheckBox.Text      = "Exclude index creation"
$excludeIndexCheckBox.Location  = New-Object System.Drawing.Point(12, 70)
$excludeIndexCheckBox.AutoSize  = $true
$excludeIndexCheckBox.ForeColor = $fgPrimary
$excludeIndexCheckBox.Checked   = $true

$excludeStatsCheckBox = New-Object System.Windows.Forms.CheckBox
$excludeStatsCheckBox.Text      = "Exclude stats updates"
$excludeStatsCheckBox.Location  = New-Object System.Drawing.Point(12, 95)
$excludeStatsCheckBox.AutoSize  = $true
$excludeStatsCheckBox.ForeColor = $fgPrimary
$excludeStatsCheckBox.Checked   = $true

$topRowsLabel = New-Object System.Windows.Forms.Label
$topRowsLabel.Text      = "Top rows:"
$topRowsLabel.Location  = New-Object System.Drawing.Point(12, 130)
$topRowsLabel.AutoSize  = $true
$topRowsLabel.ForeColor = $fgPrimary

$topNumeric = New-Object System.Windows.Forms.NumericUpDown
$topNumeric.Location    = New-Object System.Drawing.Point(120, 127)
$topNumeric.Size        = New-Object System.Drawing.Size(80, 20)
$topNumeric.Minimum     = 10
$topNumeric.Maximum     = 10000
$topNumeric.Value       = 100
$topNumeric.BackColor   = $gridRowBg
$topNumeric.ForeColor   = $fgPrimary
$topNumeric.BorderStyle = 'FixedSingle'

$loadPlansButton = New-Object System.Windows.Forms.Button
$loadPlansButton.Text      = "Run..."
$loadPlansButton.Location  = New-Object System.Drawing.Point(12, 170)
$loadPlansButton.Size      = New-Object System.Drawing.Size(268, 30)
$loadPlansButton.BackColor = $accentBlue
$loadPlansButton.ForeColor = $fgAccent
$loadPlansButton.FlatStyle = 'Flat'

$insightsButton = New-Object System.Windows.Forms.Button
$insightsButton.Text      = "Insights"
$insightsButton.Location  = New-Object System.Drawing.Point(12, 210)
$insightsButton.Size      = New-Object System.Drawing.Size(120, 30)
$insightsButton.BackColor = $accentGreen
$insightsButton.ForeColor = $fgAccent
$insightsButton.FlatStyle = 'Flat'
$insightsButton.Enabled   = $false

$saveJsonButton = New-Object System.Windows.Forms.Button
$saveJsonButton.Text      = "Save"
$saveJsonButton.Location  = New-Object System.Drawing.Point(140, 210)
$saveJsonButton.Size      = New-Object System.Drawing.Size(70, 30)
$saveJsonButton.BackColor = $btnGray
$saveJsonButton.ForeColor = $fgPrimary
$saveJsonButton.FlatStyle = 'Flat'

$loadJsonButton = New-Object System.Windows.Forms.Button
$loadJsonButton.Text      = "Load"
$loadJsonButton.Location  = New-Object System.Drawing.Point(220, 210)
$loadJsonButton.Size      = New-Object System.Drawing.Size(60, 30)
$loadJsonButton.BackColor = $btnGray
$loadJsonButton.ForeColor = $fgPrimary
$loadJsonButton.FlatStyle = 'Flat'

$queryGroupBox.Controls.AddRange(@(
    $timeWindowLabel, $timeWindowCombo,
    $excludeIndexCheckBox, $excludeStatsCheckBox,
    $topRowsLabel, $topNumeric,
    $loadPlansButton,
    $insightsButton,
    $saveJsonButton,
    $loadJsonButton
))

# ====================================================================================
# CONSOLE LOG
# ====================================================================================

$consoleGroupBox = New-DarkGroupBox -text "Console Log" `
    -location ([System.Drawing.Point]::new(340, 40)) `
    -size     ([System.Drawing.Size]::new(780, 150)) `
    -anchor   "Top, Left, Right"

$mainForm.Controls.Add($consoleGroupBox)

$consoleLogTextBox = New-Object System.Windows.Forms.TextBox
$consoleLogTextBox.Location   = New-Object System.Drawing.Point(10, 20)
$consoleLogTextBox.Size       = New-Object System.Drawing.Size(760, 120)
$consoleLogTextBox.Multiline  = $true
$consoleLogTextBox.ScrollBars = 'Vertical'
$consoleLogTextBox.ReadOnly   = $true
$consoleLogTextBox.BackColor  = $consoleBack
$consoleLogTextBox.ForeColor  = $consoleFore
$consoleLogTextBox.BorderStyle = 'FixedSingle'
$consoleLogTextBox.Anchor     = "Top, Left, Right"

$consoleGroupBox.Controls.Add($consoleLogTextBox)

# ====================================================================================
# RESULTS GRID
# ====================================================================================

$resultsGroupBox = New-DarkGroupBox -text "Query Plan Results" `
    -location ([System.Drawing.Point]::new(340, 200)) `
    -size     ([System.Drawing.Size]::new(780, 440)) `
    -anchor   "Top, Bottom, Left, Right"

$mainForm.Controls.Add($resultsGroupBox)

$plansGrid = New-Object System.Windows.Forms.DataGridView
$plansGrid.Dock                 = 'Fill'
$plansGrid.AutoSizeColumnsMode  = 'None'
$plansGrid.AllowUserToResizeColumns = $true
Apply-DarkModeToGrid -grid $plansGrid
$resultsGroupBox.Controls.Add($plansGrid)

# Context menu
$gridContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuCopyRow     = New-Object System.Windows.Forms.ToolStripMenuItem("Copy Row to Clipboard")
$menuViewPlan    = New-Object System.Windows.Forms.ToolStripMenuItem("View Execution Plan")
[void]$gridContextMenu.Items.AddRange(@($menuCopyRow, $menuViewPlan))
$plansGrid.ContextMenuStrip = $gridContextMenu

# ====================================================================================
# PLAN ROW HIGHLIGHTING
# ====================================================================================

function Highlight-PlanRows {
    param([System.Windows.Forms.DataGridView]$grid)

    $styleCritical  = New-Object System.Windows.Forms.DataGridViewCellStyle
    $styleCritical.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#7f1d1d")
    $styleCritical.ForeColor          = $fgAccent
    $styleCritical.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#991b1b")
    $styleCritical.SelectionForeColor = $fgAccent

    $styleSpill  = New-Object System.Windows.Forms.DataGridViewCellStyle
    $styleSpill.BackColor          = $blockingRed
    $styleSpill.ForeColor          = $fgAccent
    $styleSpill.SelectionBackColor = $blockingRed
    $styleSpill.SelectionForeColor = $fgAccent

    $styleImplicit  = New-Object System.Windows.Forms.DataGridViewCellStyle
    $styleImplicit.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#7c2d12")
    $styleImplicit.ForeColor          = $fgAccent
    $styleImplicit.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#9a3412")
    $styleImplicit.SelectionForeColor = $fgAccent

    $styleMissing  = New-Object System.Windows.Forms.DataGridViewCellStyle
    $styleMissing.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#064e3b")
    $styleMissing.ForeColor          = $fgAccent
    $styleMissing.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#047857")
    $styleMissing.SelectionForeColor = $fgAccent

    $styleWarn  = New-Object System.Windows.Forms.DataGridViewCellStyle
    $styleWarn.BackColor          = [System.Drawing.ColorTranslator]::FromHtml("#4c1d95")
    $styleWarn.ForeColor          = $fgAccent
    $styleWarn.SelectionBackColor = [System.Drawing.ColorTranslator]::FromHtml("#6d28d9")
    $styleWarn.SelectionForeColor = $fgAccent

    foreach ($row in $grid.Rows) {
        $data = $row.DataBoundItem
        if ($null -eq $data) { continue }

        $insights = ""
        try { $insights = [string]$data.plan_insights } catch {}

        if ([string]::IsNullOrWhiteSpace($insights)) { continue }

        if ($insights -like '*NO JOIN PREDICATE*' -or
            $insights -like '*CARDINALITY ESTIMATE WARNING*' -or
            $insights -like '*ROW GOAL*') {
            $row.DefaultCellStyle = $styleCritical
        }
        elseif ($insights -like '*SPILL TO TEMPDB*' -or
                $insights -like '*MEMORY GRANT ISSUE*') {
            $row.DefaultCellStyle = $styleSpill
        }
        elseif ($insights -like '*MISSING INDEXES*' -or
                $insights -like '*MISSING OR STALE STATS*') {
            $row.DefaultCellStyle = $styleMissing
        }
        elseif ($insights -like '*IMPLICIT CONVERSION*' -or
                $insights -like '*PLAN AFFECTING CONVERT*') {
            $row.DefaultCellStyle = $styleImplicit
        }
        elseif ($insights -like '*OTHER WARNINGS*' -or
                $insights -like '*NON PARALLEL PLAN*') {
            $row.DefaultCellStyle = $styleWarn
        }
    }
}

$plansGrid.add_DataBindingComplete({
    param($sender, $e)

    $grid = $sender

    foreach ($col in $grid.Columns) {
        if ($col.Name -eq 'query_plan') {
            $col.Visible = $false
        }
        elseif ($col.Name -eq 'plan_insights') {
            $col.HeaderText   = 'Plan Insights'
            $col.AutoSizeMode = 'None'
            $col.Width        = 260
            $col.MinimumWidth = 150
            $col.Resizable    = 'True'
        }
        elseif ($col.Name -like 'has_*') {
            $col.Visible = $false
        }
        elseif ($col.Name -eq 'query_sql_text') {
            $col.AutoSizeMode = 'None'
            $col.Width        = 450
        }
        elseif ($col.Name -eq 'last_duration_hhmmss') {
            $col.AutoSizeMode = 'None'
            $col.Width        = 95
        }
        else {
            $col.AutoSizeMode = 'AllCells'
        }
    }

    if ($grid.Columns['last_duration_sec'])      { $grid.Columns['last_duration_sec'].DisplayIndex      = 0 }
    if ($grid.Columns['last_duration_hhmmss'])   { $grid.Columns['last_duration_hhmmss'].DisplayIndex   = 1 }
    if ($grid.Columns['plan_insights'])          { $grid.Columns['plan_insights'].DisplayIndex          = 2 }
    if ($grid.Columns['query_hash_hex'])         { $grid.Columns['query_hash_hex'].DisplayIndex         = 6 }
    if ($grid.Columns['query_plan_hash_hex'])    { $grid.Columns['query_plan_hash_hex'].DisplayIndex    = 5 }
    if ($grid.Columns['query_id'])               { $grid.Columns['query_id'].DisplayIndex               = 4 }
    if ($grid.Columns['query_sql_text'])         { $grid.Columns['query_sql_text'].DisplayIndex         = 3 }

    if ($grid.Columns['query_plan_hash_hex']) {
        $hasPlanHashValue = $false
        foreach ($row in $grid.Rows) {
            $data = $row.DataBoundItem
            if ($null -eq $data) { continue }
            $val = ""
            try { $val = [string]$data.query_plan_hash_hex } catch {}
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $hasPlanHashValue = $true
                break
            }
        }
        $grid.Columns['query_plan_hash_hex'].Visible = $hasPlanHashValue
    }

    Highlight-PlanRows -grid $grid
})

# ====================================================================================
# GRID CONTEXT MENU HANDLERS
# ====================================================================================

function Get-SelectedRowObjectFromGrid {
    param([System.Windows.Forms.DataGridView]$grid)

    if ($grid.SelectedRows.Count -gt 0) {
        return $grid.SelectedRows[0].DataBoundItem
    }
    elseif ($grid.CurrentRow -and $grid.CurrentRow.Index -ge 0) {
        return $grid.CurrentRow.DataBoundItem
    }
    return $null
}

function Copy-RowToClipboard {
    param([System.Windows.Forms.DataGridView]$grid)

    $rowObj = Get-SelectedRowObjectFromGrid -grid $grid
    if (-not $rowObj) { return }

    $sbHeader = New-Object System.Text.StringBuilder
    $sbValues = New-Object System.Text.StringBuilder

    foreach ($col in $grid.Columns) {
        if ($col.Visible) {
            [void]$sbHeader.Append($col.HeaderText)
            [void]$sbHeader.Append("`t")

            $val = ""
            try { $val = $rowObj."$($col.DataPropertyName)" } catch {}

            [void]$sbValues.Append($val)
            [void]$sbValues.Append("`t")
        }
    }

    $text = ($sbHeader.ToString().Trim("`t") + "`r`n" + $sbValues.ToString().Trim("`t"))
    [System.Windows.Forms.Clipboard]::SetText($text)
    Log-Message "Row copied to clipboard."
}

function Show-ExecutionPlan {
    param([System.Windows.Forms.DataGridView]$grid)

    $rowObj = Get-SelectedRowObjectFromGrid -grid $grid
    if (-not $rowObj) { return }

    $planXml = $null
    try { $planXml = $rowObj.query_plan } catch { $planXml = $null }

    if (-not $planXml -or [string]::IsNullOrWhiteSpace([string]$planXml)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No execution plan available for this row.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $tempFile = [System.IO.Path]::Combine(
        [System.IO.Path]::GetTempPath(),
        ("QueryPlan_{0}.sqlplan" -f ([Guid]::NewGuid().ToString("N")))
    )

    try {
        [System.IO.File]::WriteAllText($tempFile, [string]$planXml, [System.Text.Encoding]::UTF8)
        Start-Process $tempFile
        Log-Message "Execution plan written to $tempFile"
    }
    catch {
        Log-Message "ERROR: Could not open execution plan. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Could not open execution plan viewer on this machine.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

$menuCopyRow.Add_Click({
    $src = $gridContextMenu.SourceControl
    if ($src -is [System.Windows.Forms.DataGridView]) {
        Copy-RowToClipboard -grid $src
    }
})

$menuViewPlan.Add_Click({
    $src = $gridContextMenu.SourceControl
    if ($src -is [System.Windows.Forms.DataGridView]) {
        Show-ExecutionPlan -grid $src
    }
})

# ====================================================================================
# SAVE / LOAD JSON
# ====================================================================================

function Save-ResultsToJson {
    if (-not $plansGrid.DataSource -or $plansGrid.DataSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no results to save.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter       = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dlg.Title        = "Save QS Insights results"
    $dlg.DefaultExt   = "json"
    $dlg.AddExtension = $true

    if ($dlg.ShowDialog($mainForm) -ne 'OK') { return }

    try {
        $data = $plansGrid.DataSource
        $json = $data | ConvertTo-Json -Depth 6
        [System.IO.File]::WriteAllText($dlg.FileName, $json, [System.Text.Encoding]::UTF8)
        Log-Message "Results saved to JSON: $($dlg.FileName)"
    }
    catch {
        Log-Message "ERROR: Failed to save JSON. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to save JSON.`r`n$($_.Exception.Message)",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Load-ResultsFromJson {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $dlg.Title  = "Load QS Insights results"

    if ($dlg.ShowDialog($mainForm) -ne 'OK') { return }

    try {
        $json = [System.IO.File]::ReadAllText($dlg.FileName, [System.Text.Encoding]::UTF8)
        $objs = $json | ConvertFrom-Json

        $list = New-Object System.Collections.ArrayList
        if ($objs -is [System.Collections.IEnumerable] -and -not ($objs -is [string])) {
            foreach ($o in $objs) { [void]$list.Add($o) }
        } else {
            [void]$list.Add($objs)
        }

        $plansGrid.DataSource = $null
        $plansGrid.DataSource = $list

        Log-Message "Results loaded from JSON: $($dlg.FileName)"

        if ($list.Count -gt 0) {
            $insightsButton.Enabled = $true
        }
    }
    catch {
        Log-Message "ERROR: Failed to load JSON. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load JSON.`r`n$($_.Exception.Message)",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

$saveJsonButton.Add_Click({ Save-ResultsToJson })
$loadJsonButton.Add_Click({ Load-ResultsFromJson })

# ====================================================================================
# CONNECTION TEST AND SESSION UI
# ====================================================================================

function Invoke-ConnectionTest {
    $server   = $serverTextBox.Text
    $username = $usernameTextBox.Text
    $password = Get-PasswordFromTextBox

    if ([string]::IsNullOrWhiteSpace($server) -or
        [string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter server and username.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $cs   = "Server=$server;Database=master;User ID=$username;Password=$password;Connection Timeout=5;TrustServerCertificate=True"
    $conn = New-Object System.Data.SqlClient.SqlConnection($cs)

    try {
        Log-Message "Testing connection to $server..."
        $conn.Open()
        $connectionStatusLabel.Text      = "Status: Connection Successful"
        $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#32CD32")
        Log-Message "Connection to $server successful."

        Load-QueryStoreDatabases -Server $server -Username $username -Password $password
    }
    catch {
        $connectionStatusLabel.Text      = "Status: Connection Failed"
        $connectionStatusLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FF4500")
        Log-Message "ERROR: Failed to connect to $server. $($_.Exception.Message)"
        $databaseComboBox.Items.Clear()
        $databaseComboBox.Enabled = $false
    }
    finally {
        if ($conn.State -eq 'Open') { $conn.Close() }
    }
}

function Populate-SessionDropdown {
    Ensure-SessionsArrayList
    $sessionComboBox.Items.Clear()

    foreach ($s in $script:Sessions) {
        [void]$sessionComboBox.Items.Add($s.Name)
    }

    if ($sessionComboBox.Items.Count -gt 0) {
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentSession) -and
            $sessionComboBox.Items.Contains($script:CurrentSession)) {
            $sessionComboBox.SelectedItem = $script:CurrentSession
        } else {
            $sessionComboBox.SelectedIndex = 0
        }
    }
}

function Set-UiFromSession {
    param([pscustomobject]$session)

    if (-not $session) { return }

    $serverTextBox.Text   = $session.Server
    $usernameTextBox.Text = $session.Username
    $passwordTextBox.Tag  = $session.Password
    $passwordTextBox.Text = '••••••••••'

    $connectionStatusLabel.Text      = "Status: Not Connected"
    $connectionStatusLabel.ForeColor = $fgSecondary

    $databaseComboBox.Items.Clear()
    $databaseComboBox.Enabled = $false
}

function Load-LastSessionToUi {
    if ([string]::IsNullOrWhiteSpace($script:CurrentSession)) { return }
    $sess = Get-SessionByName -Name $script:CurrentSession
    if ($sess) {
        Set-UiFromSession -session $sess
        $currentSessionValue.Text = $sess.Name
        if ($sessionComboBox.Items.Contains($sess.Name)) {
            $sessionComboBox.SelectedItem = $sess.Name
        }
        Log-Message "Loaded last session '$($sess.Name)'."
        Invoke-ConnectionTest
    }
}

function Prompt-ForSessionName {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text          = "Save Session As"
    $dlg.Size          = New-Object System.Drawing.Size(350, 150)
    $dlg.StartPosition = "CenterParent"
    $dlg.BackColor     = $bgMain
    $dlg.ForeColor     = $fgPrimary
    $dlg.Font          = $fontRegular
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox   = $false
    $dlg.MinimizeBox   = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = "Session Name:"
    $lbl.Location  = New-Object System.Drawing.Point(10, 20)
    $lbl.AutoSize  = $true
    $lbl.ForeColor = $fgPrimary

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location    = New-Object System.Drawing.Point(110, 18)
    $tb.Size        = New-Object System.Drawing.Size(210, 20)
    $tb.BackColor   = $gridRowBg
    $tb.ForeColor   = $fgPrimary
    $tb.BorderStyle = 'FixedSingle'

    if (-not [string]::IsNullOrWhiteSpace($script:CurrentSession)) {
        $tb.Text = $script:CurrentSession
    }

    $okBtn = New-Object System.Windows.Forms.Button
    $okBtn.Text      = "Save"
    $okBtn.Location  = New-Object System.Drawing.Point(160, 60)
    $okBtn.Size      = New-Object System.Drawing.Size(70, 28)
    $okBtn.BackColor = $btnGray
    $okBtn.ForeColor = $fgPrimary
    $okBtn.FlatStyle = 'Flat'

    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text      = "Cancel"
    $cancelBtn.Location  = New-Object System.Drawing.Point(245, 60)
    $cancelBtn.Size      = New-Object System.Drawing.Size(70, 28)
    $cancelBtn.BackColor = $btnGray
    $cancelBtn.ForeColor = $fgPrimary
    $cancelBtn.FlatStyle = 'Flat'

    $dlg.Controls.AddRange(@($lbl, $tb, $okBtn, $cancelBtn))

    $okBtn.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($tb.Text)) {
            $dlg.Tag          = $tb.Text
            $dlg.DialogResult = 'OK'
            $dlg.Close()
        }
    })
    $cancelBtn.Add_Click({
        $dlg.DialogResult = 'Cancel'
        $dlg.Close()
    })

    $res = $dlg.ShowDialog($mainForm)
    if ($res -eq 'OK') {
        return $dlg.Tag
    }
    return $null
}

$testConnectionButton.Add_Click({
    Invoke-ConnectionTest
})

$saveConnectionButton.Add_Click({
    $sessName = Prompt-ForSessionName
    if ([string]::IsNullOrWhiteSpace($sessName)) {
        Log-Message "Save session cancelled."
        return
    }

    $securePwd = ($passwordTextBox.Text | ConvertTo-SecureString -AsPlainText -Force)
    Upsert-Session -Name $sessName `
        -Server   $serverTextBox.Text `
        -Username $usernameTextBox.Text `
        -Password $securePwd

    $script:CurrentSession    = $sessName
    $currentSessionValue.Text = $sessName
    Log-Message "Session '$sessName' saved to $script:SessionsFilePath"

    Populate-SessionDropdown
    $sessionComboBox.SelectedItem = $sessName
})

$deleteSessionButton.Add_Click({
    $name = $script:CurrentSession
    if ([string]::IsNullOrWhiteSpace($name)) {
        Log-Message "No active session to delete."
        return
    }

    $res = [System.Windows.Forms.MessageBox]::Show(
        "Delete session '$name'?",
        "QS Insights",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
        Log-Message "Delete cancelled."
        return
    }

    Remove-SessionByName -Name $name
    Log-Message "Session '$name' deleted."

    $serverTextBox.Text   = ""
    $usernameTextBox.Text = ""
    $passwordTextBox.Text = ""
    $passwordTextBox.Tag  = $null
    $currentSessionValue.Text = "(none)"
    $script:CurrentSession    = $null
    $connectionStatusLabel.Text      = "Status: Not Connected"
    $connectionStatusLabel.ForeColor = $fgSecondary

    $databaseComboBox.Items.Clear()
    $databaseComboBox.Enabled = $false

    Populate-SessionDropdown
})

$sessionComboBox.add_SelectedIndexChanged({
    $chosenName = $sessionComboBox.SelectedItem
    if ([string]::IsNullOrWhiteSpace($chosenName)) { return }

    $sess = Get-SessionByName -Name $chosenName
    if ($sess) {
        $script:CurrentSession    = $sess.Name
        $currentSessionValue.Text = $sess.Name
        Set-UiFromSession -session $sess
        Save-AllSessions
        Log-Message "Session '$($sess.Name)' loaded from selector."
        Invoke-ConnectionTest
    }
})

# ====================================================================================
# RUN QUERY PLANS (PHASE 1, NO XML INSIGHTS)
# ====================================================================================

function Run-PlanQuery {
    $server   = $serverTextBox.Text
    $username = $usernameTextBox.Text
    $password = Get-PasswordFromTextBox
    $dbName   = [string]$databaseComboBox.SelectedItem

    if ([string]::IsNullOrWhiteSpace($server) -or
        [string]::IsNullOrWhiteSpace($username)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please provide server and username first.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($dbName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a database with Query Store enabled.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $connStr = "Server=$server;Database=$dbName;User ID=$username;Password=$password;TrustServerCertificate=True;Application Name=QS Insights"
    $top     = [int]$topNumeric.Value
    $timeExpr = Get-QueryStoreTimeWindowExpression

    $whereParts = @()
    $whereParts += "rsi.end_time > $timeExpr"

    if ($excludeIndexCheckBox.Checked) {
        $whereParts += "qt.query_sql_text NOT LIKE 'ALTER INDEX%' and qt.query_sql_text NOT LIKE 'CREATE%INDEX%'"
    }
    if ($excludeStatsCheckBox.Checked) {
        $whereParts += "qt.query_sql_text NOT LIKE 'UPDATE STATISTICS%'"
    }

    $whereClause = "WHERE " + ($whereParts -join "`r`n  AND ")

    # Detect whether query_plan_hash column exists
    $hasPlanHash = $false
    try {
        $check = Get-SqlData -ConnectionString $connStr -Query @"
IF EXISTS (
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID('sys.query_store_query')
      AND name = 'query_plan_hash'
)
    SELECT CAST(1 AS int) AS HasPlanHash;
ELSE
    SELECT CAST(0 AS int) AS HasPlanHash;
"@
        if ($check -and $check.Count -gt 0 -and [int]$check[0].HasPlanHash -eq 1) {
            $hasPlanHash = $true
        }
    }
    catch {
        Log-Message "Warning: Could not probe query_plan_hash column ($($_.Exception.Message)). Assuming it does not exist."
        $hasPlanHash = $false
    }

    if ($hasPlanHash) {
        $planHashColSql = "CONVERT(varchar(34), sys.fn_varbintohexstr(CAST(q.query_plan_hash AS varbinary(8)))) AS query_plan_hash_hex,"
        Log-Message "query_plan_hash column detected; returning plan hash."
    } else {
        $planHashColSql = "CAST(NULL AS varchar(34)) AS query_plan_hash_hex,"
        Log-Message "query_plan_hash column not present; returning NULL for plan hash."
    }

    $tsqlTemplate = @"
SELECT TOP ($top)
       rs.last_duration,                                                -- microseconds
       CONVERT(decimal(18,3), rs.last_duration / 1000000.0) AS last_duration_sec,
       CONVERT(varchar(12), DATEADD(ms, rs.last_duration / 1000, '00:00:00'), 114) AS last_duration_hhmmss,
       q.query_id,
       CONVERT(varchar(34), sys.fn_varbintohexstr(CAST(q.query_hash AS varbinary(8))))      AS query_hash_hex,
       __PLANHASH_COL__
       qt.query_sql_text,
       p.plan_id,
       p.last_execution_time,
       CAST(0 AS int) AS has_missing_index,
       CAST(0 AS int) AS has_implicit_conversion,
       CAST(0 AS int) AS has_spill,
       CAST(0 AS int) AS has_warnings,
       CAST(0 AS int) AS has_memory_grant_issue,
       CAST(0 AS int) AS has_no_join_predicate,
       CAST(0 AS int) AS has_missing_stats,
       CAST(0 AS int) AS has_plan_affecting_convert,
       CAST(0 AS int) AS has_row_goal,
       CAST(0 AS int) AS has_cardinality_estimate_warning,
       CAST(0 AS int) AS has_nonparallel_reason,
       CAST('' AS nvarchar(4000)) AS plan_insights,
       CONVERT(nvarchar(max), p.query_plan) AS query_plan
FROM sys.query_store_runtime_stats AS rs
JOIN sys.query_store_runtime_stats_interval AS rsi
    ON rs.runtime_stats_interval_id = rsi.runtime_stats_interval_id
JOIN sys.query_store_plan AS p
    ON rs.plan_id = p.plan_id
JOIN sys.query_store_query AS q
    ON p.query_id = q.query_id
JOIN sys.query_store_query_text AS qt
    ON q.query_text_id = qt.query_text_id
$whereClause
ORDER BY rs.last_duration DESC;
"@

    $tsql = $tsqlTemplate -replace '__PLANHASH_COL__', $planHashColSql

    try {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $insightsButton.Enabled = $false
        Log-Message "Running Query Store plan search on [$dbName] (Top $top) for $($timeWindowCombo.SelectedItem) (phase 1, no insights)..."

        $rows = Get-SqlData -ConnectionString $connStr -Query $tsql
        $plansGrid.DataSource = $null

        if ($rows -and $rows.Count -gt 0) {
            $plansGrid.DataSource = [System.Collections.ArrayList]$rows
            Log-Message ("Loaded {0} rows into grid. Click Insights to analyse execution plans locally." -f $rows.Count)
            $insightsButton.Enabled = $true
        }
        else {
            Log-Message "Query returned no rows."
        }
    }
    catch {
        Log-Message "ERROR: Failed to load query plans. $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load query plans.`r`n$($_.Exception.Message)",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$loadPlansButton.Add_Click({
    Run-PlanQuery
})

# ====================================================================================
# PHASE 2: LOCAL PLAN INSIGHTS FROM XML
# ====================================================================================

function Test-XmlNodePresent {
    param(
        [xml]$XmlDoc,
        [string]$XPath
    )
    try {
        return ($null -ne $XmlDoc.SelectSingleNode($XPath))
    }
    catch {
        return $false
    }
}

function Generate-LocalPlanInsights {
    if (-not $plansGrid.DataSource -or $plansGrid.DataSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "There are no results loaded. Run a query or load JSON first.",
            "QS Insights",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $data = [System.Collections.ArrayList]$plansGrid.DataSource

    $propsToEnsure = @(
        'has_missing_index',
        'has_implicit_conversion',
        'has_spill',
        'has_warnings',
        'has_memory_grant_issue',
        'has_no_join_predicate',
        'has_missing_stats',
        'has_plan_affecting_convert',
        'has_row_goal',
        'has_cardinality_estimate_warning',
        'has_nonparallel_reason',
        'plan_insights'
    )

    $mainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $insightsButton.Enabled = $false

    Log-Message "Analysing execution plans locally for insights..."

    $processed = 0

    foreach ($row in $data) {
        foreach ($pName in $propsToEnsure) {
            if (-not $row.PSObject.Properties.Match($pName)) {
                $default = 0
                if ($pName -eq 'plan_insights') { $default = '' }
                $row | Add-Member -NotePropertyName $pName -Value $default
            }
        }

        $xmlText = $null
        try { $xmlText = [string]$row.query_plan } catch { $xmlText = $null }

        if ([string]::IsNullOrWhiteSpace($xmlText)) { continue }

        $doc = $null
        try { $doc = [xml]$xmlText } catch { continue }

        $hasMissingIndex = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='MissingIndex']"
        $hasImplicitConv = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='ScalarOperator' and contains(@ScalarString,'CONVERT_IMPLICIT')]"
        $hasPlanAffConv  = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='PlanAffectingConvert']"

        # spills: any of the known spill nodes
        $hasSpill = $false
        if (Test-XmlNodePresent -XmlDoc $doc -XPath "//*[@SpillToTempDb='1' or @SpillToTempDb='true']") {
            $hasSpill = $true
        }
        elseif (Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='SortSpillDetails']") {
            $hasSpill = $true
        }
        elseif (Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='HashSpillDetails']") {
            $hasSpill = $true
        }

        $hasWarnings  = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='Warnings']/*"
        $hasMemGrant  = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='MemoryGrantWarning' and @GrantWarningKind!='None']"
        $hasNoJoin    = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='NoJoinPredicate']"
        $hasRowGoal   = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='RowGoal']"
        $hasMissStats = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='ColumnsWithNoStatistics' or local-name()='MissingStatistics']"
        $hasCEWarning = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='CardinalityEstimateWarning']"
        $hasNonParallel = Test-XmlNodePresent -XmlDoc $doc -XPath "//*[local-name()='QueryPlan' and string-length(@NonParallelPlanReason) > 0]"

        $labels = @()
        if ($hasSpill)       { $labels += 'SPILL TO TEMPDB' }
        if ($hasMemGrant)    { $labels += 'MEMORY GRANT ISSUE' }
        if ($hasImplicitConv){ $labels += 'IMPLICIT CONVERSION' }
        if ($hasPlanAffConv) { $labels += 'PLAN AFFECTING CONVERT' }
        if ($hasMissingIndex){ $labels += 'MISSING INDEXES' }
        if ($hasMissStats)   { $labels += 'MISSING OR STALE STATS' }
        if ($hasNoJoin)      { $labels += 'NO JOIN PREDICATE' }
        if ($hasRowGoal)     { $labels += 'ROW GOAL' }
        if ($hasCEWarning)   { $labels += 'CARDINALITY ESTIMATE WARNING' }
        if ($hasNonParallel) { $labels += 'NON PARALLEL PLAN' }
        if ($hasWarnings -and -not $hasCEWarning) { $labels += 'OTHER WARNINGS' }

        $row.has_missing_index                = [int]$hasMissingIndex
        $row.has_implicit_conversion          = [int]$hasImplicitConv
        $row.has_spill                        = [int]$hasSpill
        $row.has_warnings                     = [int]$hasWarnings
        $row.has_memory_grant_issue           = [int]$hasMemGrant
        $row.has_no_join_predicate            = [int]$hasNoJoin
        $row.has_missing_stats                = [int]$hasMissStats
        $row.has_plan_affecting_convert       = [int]$hasPlanAffConv
        $row.has_row_goal                     = [int]$hasRowGoal
        $row.has_cardinality_estimate_warning = [int]$hasCEWarning
        $row.has_nonparallel_reason           = [int]$hasNonParallel
        $row.plan_insights                    = ($labels -join '; ')

        $processed++
    }

    $plansGrid.DataSource = $null
    $plansGrid.DataSource = $data

    Log-Message ("Plan insights generated for {0} row(s)." -f $processed)

    $mainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    if ($data.Count -gt 0) { $insightsButton.Enabled = $true }
}

$insightsButton.Add_Click({
    Generate-LocalPlanInsights
})

# ====================================================================================
# MENU EVENTS
# ====================================================================================

$miExit.Add_Click({
    $mainForm.Close()
})

$miAbout.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "QS Insights`r`n`r`nPhase 1: pull longest running queries from Query Store.`r`nPhase 2: generate local Plan Insights from the loaded plans.`r`nYou can also save/load result sets as JSON.",
        "About QS Insights",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
})

# ====================================================================================
# FORM SHOWN
# ====================================================================================

$mainForm.Add_Shown({
    Load-AllSessions
    Populate-SessionDropdown

    if (-not [string]::IsNullOrWhiteSpace($script:CurrentSession)) {
        $currentSessionValue.Text = $script:CurrentSession
    } else {
        $currentSessionValue.Text = "(none)"
    }

    Load-LastSessionToUi
})

[void]$mainForm.ShowDialog()
$mainForm.Dispose()
