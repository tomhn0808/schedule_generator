<#
.SYNOPSIS
    Generates a schedule of time slots based on a CSV input file.

.DESCRIPTION
    Reads an input CSV file (Windows-1252 encoding), filters students by group (A to E or '/' for all),
    handles time slots and optional break periods, then outputs a CSV file with each person's
    Name, First Name, and assigned Time Slot. This program can be used with CSV formatted in English or French

.PARAMETER inputFile
    Path to the input CSV file. Required.
    The file must include at least one column named 'Nom' or 'Name' and a 'groupe' column.

.PARAMETER group
    One or more group letters to include (A, B, C, D, E) or '/' to accept all A to E groups.
    Examples: `-group A,C,E` or `-group /`
    Comparison is case-insensitive.

.PARAMETER outputFile
    Path to the output CSV file. Default is `.\schedule.csv`.
    The output will contain filtered columns and each person's time slot.

.PARAMETER slotDuration
    Duration of each time slot in minutes (1 to 240). Default is 120.

.PARAMETER nbStudentPerSlot
    Number of people per time slot. Default is 4.

.PARAMETER scheduleStart
    Start time for the first slot, in HH:mm format. Default is '08:30'.

.PARAMETER scheduleEnd
    End time for the schedule, in HH:mm format. Default is '17:00'.

.PARAMETER breakTimes
    Optional list of break periods, each in HH:mm-HH:mm format, separated by commas.
    Example: `-breakTimes '10:00-10:30','12:00-13:00'`

.EXAMPLE
    .\schedule.ps1 -inputFile students.csv -group A,B -outputFile planning.csv

    Filters people in groups A and B, then generates planning.csv with their slots.

.EXAMPLE
    .\schedule.ps1 -inputFile students.csv -group / -breakTimes '12:00-13:00'

    Accepts all groups A to E, skips the lunch break, and writes in .\schedule.csv.

.NOTES
    - Reads the input CSV as Windows-1252 and writes the output in UTF-8 with BOM.  
    - The script throws an error and exits if a person’s group is invalid.  
    - For full documentation, run `Get-Help .\schedule.ps1 -Full` or read on the Github page.
#>
# Parameter
param (
    [Parameter(Mandatory = $true,
        HelpMessage = "Please enter the input file path",
        Position = 0)]
    [System.IO.FileInfo]$inputFile,

    # Group associated list (for cleaning up later) (separated by comma (,))
    [Parameter(Mandatory = $true,
        HelpMessage = "Enter the group(s) letter(s) or '/'",
        Position = 1)]
    [ValidateSet('A','B','C','D','E','/', IgnoreCase=$true)]
    [string[]]$group,

    # Output file containing schedule
    [Parameter(Position = 2)]
    [System.IO.FileInfo]$outputFile = ".\schedule.csv",

    # Slot duration (in minutes)
    [Parameter(Position = 3)]
    [ValidateRange(1, 240)] # Maximum 240 minutes for a slot (4h)
    [int]$slotDuration = 120,

    # Number of people per slot
    [Parameter(Position = 4)]
    [int]$nbStudentPerSlot = 4,

    # schedule start time
    [Parameter(Position = 5, Mandatory = $false)]
    [string]$scheduleStart = "08:30",

    # end schedule time
    [Parameter(Position = 6, Mandatory = $false)]
    [string]$scheduleEnd = "17:00",

    # break time list (separated by comma (,))
    [Parameter(Position = 7, Mandatory = $false)]
    [ValidateScript({
            foreach ($time in $_) {
                if ($time -notmatch '^[0-9]{2}:[0-9]{2}-[0-9]{2}:[0-9]{2}$') {
                    throw "Invalid format: '$time'. Use HH:mm-HH:mm, e.g. : 10:00-12:00"
                }
            }
            return $true
        })]
    [string[]]$breakTimes = ""
)

$tmpPath = ".\cleaned_temp.csv"

if (-not(Test-Path $inputFile)) {
    Write-Error "The input file: $inputFile does not exist"
    exit 1
}
if ($scheduleStart -notmatch '^[0-2][0-9]:[0-5][0-9]$') {
    throw "Invalid format: '$scheduleStart'. Use HH:mm (e.g., 08:30)"
}
if ($scheduleEnd -notmatch '^[0-2][0-9]:[0-5][0-9]$') {
    throw "Invalid format: '$scheduleStart'. Use HH:mm (e.g., 08:30)"
}
#[datetime]$scheduleStart=$scheduleStart
#[datetime]$scheduleEnd=$scheduleEnd

# Transform line to be Latin1 encoded (Windows-1252)
$bytes = [System.IO.File]::ReadAllBytes($inputFile)
$encoding = [System.Text.Encoding]::GetEncoding(1252)
$fullText = $encoding.GetString($bytes)
$rawLines = $fullText -split "`r?`n"

$startLine = ($rawLines | Select-String "Name|Nom").LineNumber - 1
$cleanLines = $rawLines[$startLine..($rawLines.Length - 1)] | Where-Object {
    $line = $_.Trim()
    # Ignore empty or whitespace-only lines
    if ($line -eq "") { return $false }

    # Check if all fields are empty after splitting on "," or ";"
    $fields = $line -split '[,;]'
    return $fields -ne ""  # Keep if at least one non-empty field
}

$normalizedLines = $cleanLines -replace ';;', ';'
$normalizedLines = $normalizedLines.ToLower()
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllLines($tmpPath, $normalizedLines, $utf8Bom)

$students = Import-Csv -Path $tmpPath -Delimiter ';' | Where-Object { ($_.Nom -ne "Nom") -or ($_.Name -ne "Name") } 

if ($students.Count -eq 0) {
    Write-Error "No students have been imported. Verify file syntax."
    exit 2
}
$group = $group | ForEach-Object { $_.ToUpper() }


# Determine which property name to use
if ($students[0].PSObject.Properties.Name -contains 'groupe') {
    $groupProp='groupe'
}
elseif ($students[0].PSObject.Properties.Name -contains 'group') {
    $groupProp='group'
}
else {
    throw "Neither 'groupe' nor 'group' column found in CSV."
}

# Validate group, else throw student
$students = $students | Where-Object {
    $studentGroup = ( $_.$groupProp ).ToUpper()
    $letterInGroup = $null

    if ($studentGroup -match '[A-E]|/') {
        $letterInGroup = $matches[0]
    }

    if (-not $letterInGroup) {
        Write-Host "Student '$($_.Nom)' has no valid group letter A-E or / in '$studentGroup'" -ForegroundColor Red
        return $false  # Exclude student
    }
    # Accept all valid letters if "/" is in $group
    if ($group -contains "/") {
        return $true
    }

    if (-not ($group -contains $letterInGroup)) {
        Write-Host "Student '$($_.Nom)' is in the wrong group ('$letterInGroup'). Not allowed here" -ForegroundColor Red
        return $false  # Exclude student
    }

    return $true  # Keep student
}

# Shuffle students
$shuffled = $students | Get-Random -Count $students.Count

# Schedule parameters
$slotDurationCleaned = [timespan]::FromMinutes($slotDuration)
$scheduleStart = [datetime]::ParseExact($scheduleStart, "HH:mm", $null)
$scheduleEnd = [datetime]::ParseExact($scheduleEnd, "HH:mm", $null)

# Create datetime interval for break time
$parsedBreaks = foreach ($b in $breakTimes) {
    $parts = $b -split '-'
    if ($parts.Count -eq 2) {
        [PSCustomObject]@{
            Start = [datetime]::ParseExact($parts[0], 'HH:mm', $null)
            End   = [datetime]::ParseExact($parts[1], 'HH:mm', $null)
        }
    }
}

$schedule = @()
$currentTime = $scheduleStart
$studentIndex = 0

while ($studentIndex -lt $shuffled.Count -and $currentTime -lt $scheduleEnd) {

    # Verify if time slot in a break time
    $inBreak = $false
    foreach ($b in $parsedBreaks) {
        if ($currentTime -ge $b.Start -and $currentTime -lt $b.End) {
            $inBreak = $true
            $currentTime = $b.End  # jump to the end of the break
            break
        }
    }

    if ($inBreak) { continue }

    # Put student to time slot
    for ($i = 0; $i -lt $nbStudentPerSlot -and $studentIndex -lt $shuffled.Count; $i++) {
        $timeSlot = "{0:HH:mm}" -f $currentTime
        $shuffled[$studentIndex] | Add-Member -NotePropertyName "Horaire" -NotePropertyValue $timeSlot
        $schedule += $shuffled[$studentIndex]
        $studentIndex++
    }

    # Go to next time slot
    $currentTime = $currentTime.AddMinutes($slotDuration)
}

$schedule
$scheduleCleaned = $schedule | Select-Object -Property * -ExcludeProperty mail


# Write clean schedule to CSV
try {
    $scheduleCleaned | Export-Csv -Path $outputFile -NoTypeInformation -Delimiter ';' -Encoding UTF8
}
catch {
    Write-Error "There has been an error writing schedule to $outputFile"
    exit 3
}
Write-Host "Planning generated" -ForegroundColor Green