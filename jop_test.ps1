$env={

    Function A2($text1, $text2)
    {
         $s=Get-Date

         $start_n = 10000000
         $end_n = $start_n *2
         $rand = Get-Random -Maximum $end_n -Minimum $start_n

         foreach ($number in 1..$rand){
            $result = $result * $number
        }

        $e=Get-Date
        $time = ($e - $s).TotalSeconds

        return @("$text1 $text2", "$time sec")
    }
}

$output = ".\output.csv"
Function tocsv($par1, $par2="", [System.ConsoleColor]$color)
{
    $properties = [PSCustomObject]@{
                time= Get-Date
                par1=$par1 
                par2=$par2
                }

    Write-Host -NoNewline $par1 
    Write-Host ", " $par2 -ForegroundColor $color
    $properties | Export-Csv $output -Append -NoTypeInformation -Encoding UTF8 
        
}


$initScript = $env
$wert = "Job"

$maxJobs = 10
$maxRunTime = 7
$global:jobs = @{}

$s=Get-Date

function myJobHandler()
{
    $doneJob = Get-Job | Where-Object { $_.State -eq 'Completed' } | Select-Object -First 1
    if ($doneJob) {
        $erg = Receive-Job -Job $doneJob
        tocsv $erg[0] $erg[1] -color Green
        Remove-Job -Job $doneJob
        $global:jobs.Remove($doneJob.Id)

        return
    } 

    $time_now = Get-Date
    $overcooked = $($global:jobs.GetEnumerator() | Sort-Object {$_.Value.time} | Select-Object -First 1 )
    $dTime = $($time_now - $overcooked.Value.time).TotalSeconds

    if ($dTime -ge $maxRunTime){
        Stop-Job -Id $overcooked.Name #Name ist Id :)
        Remove-Job -Id $overcooked.Name
        tocsv "Job $($overcooked.value.index)" "gestopt nach overtime $dTime s" -color red
        $global:jobs.Remove($overcooked.Name)

        return
    }

    Start-Sleep -Seconds 1
    Write-Host "."

}

#------------------------- Start -----------------------

Get-Job | Remove-Job

foreach($index in 1..50){

    while ($global:jobs.Count -ge $maxJobs) {
        myJobHandler
    }


    $script = 
    {
        param($para1, $para2)
        A2($para2, $para1)
    }

    #Start-Job
    Write-Host "Start $index"
    $jop = Start-Job -InitializationScript $initScript -ScriptBlock $script -ArgumentList ($index, $wert)

    $data = [PSCustomObject]@{
                time= Get-Date
                index=$index
                }
    $global:jobs.Add( $jop.Id, $data)
}

while ($global:jobs.Count -gt 0) {
    myJobHandler
}


$e=Get-Date
$time = ($e - $s).TotalSeconds
Write-Host "runntime $time"
