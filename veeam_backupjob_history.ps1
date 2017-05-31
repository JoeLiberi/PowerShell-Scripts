
if ( (Get-PSSnapin -Name veeampssnapin -ErrorAction SilentlyContinue) -eq $null )
{
    Add-PsSnapin veeampssnapin
}

$job_name = Read-Host -Prompt 'Backup Job Name'
$days = Read-Host -Prompt 'Number of days to fetch'
$folder = Read-Host -Prompt 'Enter the full path where you would like to save the output. FOLDER ONLY. I will create the file. dont forget the trailing "\"'

$Date = Get-Date 
$Filename = $folder + "_" + "$env:computername" + "_" + $Date.Month + "-" + $Date.Day + "-" + $Date.Year + "-" + $Date.Hour + "-" + $Date.Minute + "-" + $Date.Second + ".csv"

$Jobs = Get-VBRJob | ?{$_.Name -match $job_name}
$sessionstofetch = $days
$report = @()
foreach ($job in $Jobs){
	
	$jobName = $job.Name
    $table = New-Object system.Data.DataTable "$table01"
	
    $col1 = New-Object system.Data.DataColumn Index,([int])
    $col2 = New-Object system.Data.DataColumn JobName,([string])
    $col3 = New-Object system.Data.DataColumn StartTime,([DateTime])
    $col4 = New-Object system.Data.DataColumn StopTime,([DateTime])
    $col5 = New-Object system.Data.DataColumn FileName,([string])
    $col6 = New-Object system.Data.DataColumn CreationTime,([DateTime])
    $col7 = New-Object system.Data.DataColumn AvgSpeedMB,([int])
    $col8 = New-Object system.Data.DataColumn Duration,([TimeSpan])
    $col9 = New-Object system.Data.DataColumn Result,([String])

    $table.columns.add($col1)
    $table.columns.add($col2)
    $table.columns.add($col3)
    $table.columns.add($col4)
    $table.columns.add($col5)
    $table.columns.add($col6)
    $table.columns.add($col7)
    $table.columns.add($col8)
    $table.columns.add($col9)

   $session = Get-VBRBackupSession | ?{$_.JobId -eq $job.Id} | %{
		$row = $table.NewRow()
		$row.JobName = $_.Info.JobName
		$row.StartTime = $_.Info.CreationTime
		$row.StopTime = $_.Info.EndTime
		#Work out average speed in MB and round this to 0 decimal places, just like the Veeam GUI does.
		$row.AvgSpeedMB = [Math]::Round($_.Info.Progress.AvgSpeed/1024/1024,0) 
		#Duration is a Timespan value, so I am formatting in here using 3 properties - HH,MM,SS
		$row.Duration = '{0:00}:{1:00}:{2:00}' -f $_.Info.Progress.Duration.Hours, $_.Info.Progress.Duration.Minutes, $_.Info.Progress.Duration.Seconds
		$row.Result = $_.Info.Result
		
		#Add this calculated row to the $table.Rows
		$table.Rows.Add($row)
		
	}
	
	$interestingsess = $table | Sort StartTime -descending | select -first $sessionstofetch
    $pkc = 1
    $interestingsess | foreach {
		#for every object in $interestingsess (which has now been sorted by StartTime) assign the current value of $pkc to the .Index property. 1,2,3,4,5,6 etc...
		$_.Index = $pkc 
		#Increment $pkc, so the next foreach loop assigns a higher value to the next .Index property on the next row.
		$pkc+=1
	}
	
	#Now we are grabbing all the backup objects (same as viewing Backups in the Veeam B&R GUI Console
	$backup = Get-VBRBackup | ?{$_.JobId -eq $job.Id}
    $points = $backup | sort CreationTime -descending | Select -First $sessionstofetch #Find and assign the Veeam Backup files for each job we are going through and sort them in descending order. Select the specified amount.
    #Increment variable is set to 1 to start off
	$ic = 1 
    ForEach ($point in $points) {
       #Match the $ic (Increment variable) up with the Index number we kept earlier, and assign $table to $rows where they are the same. This happens for each object in $points
	   $rows = $table | ?{$_.Index -eq $ic}
       #inner ForEach loop to assign the value of the backup point's filename to the row's .FileName property as well as the creation time.
	   ForEach ($row in $rows) { 
          ($row.FileName = $point.FileName) -and ($row.CreationTime = $point.CreationTime) 
		  #Increment the $ic variable ( +1 )
		  $ic+=1
		  
       }
    }
	#Tally up the current results into our $Report Array (add them)
    $report += $interestingsess 
    
}
$report = $report | Select JobName, StartTime, StopTime, FileName, CreationTime, AvgSpeedMB, Duration, Result
$report | Export-Csv $Filename -NoTypeInformation