#Github File

function Get-MOODKDriveSpace {
    $results = @()

    $server = 'MOOCPHLDC'
    $drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
    foreach ($drive in $drives) {
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
        }
    }

    $server = 'MOOCPHLDCI'
    $drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
    foreach ($drive in $drives) {
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
        }
    }

    #Clipboard

$server = 'MOOCPHLDC0'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCRDGW1'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCRDGW2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCRSH1'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCRDSH2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCRDSH3'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHLV2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCMAIL'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOODKBACKUP'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'LEPIDE'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL3'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL4'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL8'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL11'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHL15'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPHLFORMS'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

$server = 'MOOCPROLAWMAIN'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    $results += [PSCustomObject]@{
        Server = $server
        DeviceID = $drive.DeviceID
        "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
        "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
        "Drive Free pct" = "{0,6:P0}" -f(($drive.freespace/1gb) / ($drive.size/1gb))
    }
}

    return $results
}

function ConvertTo-HTMLTable {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [PSCustomObject]$InputObject
    )
    Begin {
        $html = "<table>`n"
        $headers = $false
    }
    Process {
        if (-not $headers) {
            $html += "<tr>`n"
            foreach ($property in $InputObject.PSObject.Properties) {
                $html += "<th>$($property.Name)</th>`n"
            }
            $html += "</tr>`n"
            $headers = $true
        }
        if ($InputObject.PSObject.Properties['Drive Free Space(GB)']) {
            if ($InputObject.PSObject.Properties['Drive Free Space(GB)'].Value -lt 25) {
                $html += "<tr style='background-color: red;'>`n"
            } else {
                $html += "<tr style='background-color: lightgreen;'>`n"
            }
        } else {
            $html += "<tr>`n"
        }
        foreach ($property in $InputObject.PSObject.Properties) {
            $html += "<td>$($property.Value)</td>`n"
        }
        $html += "</tr>`n"
    }
    End {
        $html += "</table>"
        Write-Output $html
    }
}


$htmlbody = Get-MOODKDriveSpace | ConvertTo-HTMLTable 

# Send email with HTML tables
Send-MailMessage -From 'RScharf <rscharf@moodklaw.com>' -To 'RScharf <rscharf@moodklaw.com>' -Subject 'Current Drive Space - All Servers' -Body "$htmlBody" -BodyAsHtml -SmtpServer exchange.moodklaw.com
