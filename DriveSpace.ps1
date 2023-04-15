#Drive Space
#JBlack 4/15/2023, 2:52PM

#This segment establishes the Get-MoodkDriveSpace function as a drive space query to the below servers. 
function Get-MOODKDriveSpace {
    $results = @()

$server = 'MOOCPHLDC'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHLDCI'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHLDC0'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCFS'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHLV2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCRDGW1'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCRDGW2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCRDSH1'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCRDSH2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCRDSH3'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}


$server = 'MOOCMAIL'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOODKBACKUP'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'LEPIDE'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL2'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL3'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL4'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL8'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL11'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHL15'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPHLFORMS'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

$server = 'MOOCPROLAWMAIN'
$drives = Get-WmiObject Win32_logicaldisk -ComputerName $server
foreach ($drive in $drives) {
    if ($null -ne $drive.size) {
        if ($drive.size -eq 0) {
            $freeSpacePct = "0%"
        } else {
            $freeSpacePct = "{0,6:P0}" -f (($drive.freespace/1gb) / ($drive.size/1gb))
        }
    
        $results += [PSCustomObject]@{
            Server = $server
            DeviceID = $drive.DeviceID
            "Drive Size(GB)" = [decimal]("{0:N0}" -f($drive.size/1gb))
            "Drive Free Space(GB)" = [decimal]("{0:N0}" -f($drive.freespace/1gb))
            "Drive Free pct" = $freeSpacePct
        }
    }

}

    return $results
}

#This segment defines the ConverTo-HTMLTable function. It takes a pscustomobject and converts it to a HTML table with headers
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
        if ($InputObject.PSObject.Properties['Drive Free pct']) {
            if ([double]::Parse($InputObject.PSObject.Properties['Drive Free pct'].Value.TrimEnd('%')) -lt 25) {
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

#This line converts the output of Get-MOODKDriveSpace to a HTML table
$htmlbody = Get-MOODKDriveSpace | ConvertTo-HTMLTable 

# This line sends the contents of $htmlbody from my account to the Information Systems distribution list
Send-MailMessage -From 'RScharf <rscharf@moodklaw.com>' -To 'Information Systems <informationsystems@moodklaw.com>' -Subject 'Current Drive Space - All Servers' -Body "$htmlBody" -BodyAsHtml -SmtpServer exchange.moodklaw.com