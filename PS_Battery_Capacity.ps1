# Variables
    $filepath = "c:\Batterycapacity"
    $computer = "localhost"
    $batterypresent = Get-WmiObject -Class win32_battery -ComputerName $computer
    $date = get-date -f yyyy-MM-dd
    $user = Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username
    $manufacturer = Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty Manufacturer

# Script name
function MyCommandName() { return $MyInvocation.MyCommand.Name; }

# Check batteries and export data to xml
If ($batterypresent -ne $null) {
        # Check the folder, if doesn't exist create it 
        if (Get-ChildItem $filepath){
            write-host "The folder is already exist"
            } else {
                Write-host "Create folder"
                mkdir $filepath
            }
        # Make Script log (scriptname, computername, date)
Start-Transcript -path "$filepath\log\log_$($MyInvocation.MyCommand.Name)_$($env:computername)_$(get-date -f "yyyy-MM-dd_HH-mm-ss").txt"


        powercfg /batteryreport /XML /OUTPUT "$filepath\$($env:computername)_batteryreport.xml"
    Start-Sleep 1
        [xml]$Report = Get-Content "$filepath\$($env:computername)_batteryreport.xml"
}
Else {exit}

# Read xml data
$BatteryStatus = $Report.BatteryReport.Batteries.ChildNodes |
ForEach-Object {
    [PSCustomObject]@{
        DesignCapacity = $_.DesignCapacity
        FullChargeCapacity = $_.FullChargeCapacity
        Id = $_.id
        CycleCount = $_.CycleCount
 }
}

# Calculate capacities
Foreach ($Battery in $BatteryStatus) {
             ($capacityresult = [int64]$($Battery.FullChargeCapacity) * 100 / [int64]$($Battery.DesignCapacity))

# If the result is greater then 100% change the FullchargeCapacity and DesignCapacity result.
If ($capacityresult -gt 100) {
        ($capacityresult = [int64]$($Battery.DesignCapacity) * 100 / [int64]$($Battery.FullChargeCapacity))
}

# Result put out to file (Username, Computername, Battery id, Capacity, Date)
"$($user) - $($env:computername): $($Battery.id) capacity: $($capacityresult) % - $($date)"  | Out-File $filepath\$($date)_$($env:computername).txt -Verbose -Append
# Foreach end
} 

# Put to file if the result is less then 40%
If($capacityresult -lt 40) {
# Critical result put out to file
"$($user) - $($env:computername) - $($Battery.id): Critical capacity result: $capacityresult %"  | Out-File $filepath\critical_battery_result_$($env:computername).txt -Encoding utf8 -Append
}

Stop-Transcript
# Delete xml
rm  -fo $filepath\$($env:computername)_batteryreport.xml

ii $filepath
