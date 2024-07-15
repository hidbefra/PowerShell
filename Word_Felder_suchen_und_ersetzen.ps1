#ERROR REPORTING ALL
Set-StrictMode -Version latest

$env={ #das sin die Funktionen welche im Job verfügbahr sind
    $Replace_List = @{}
        $Replace_List.add("PRODSTRING3", "PRODSTRING2")
        $Replace_List.add("PROFSTRING2", "PROFSTRING6")
        $Replace_List.add("PROFSTRING9", "PRODSTRING1")

    function myReplace($text){
        foreach($key in $Replace_List.Keys){
            $text = $text.replace($key, $Replace_List[$key])
        }
        return $text
    }

    function My_count_occurrence($text){
    $count = 0
    foreach($key in $Replace_List.Keys){
        $count += ([regex]::Matches($text, $key )).count
    }
    return $count
    }

    Function process_Word($path){
        $application = New-Object -comobject word.application
        $application.visible = $true

        try{
            $document = $application.documents.open($path,$false,$false) #read and write
            } catch {
                $application.quit()
                Write-Host "$path fehlerbehandung beim öffnen" -ForegroundColor DarkMagenta
                return @($false, $path, "kann nicht geöffnet werden", [System.ConsoleColor]::Red)
            
            }
        $range = $document.content
        $count = 0

        #----- Suchen nach Felder
        #Dokumnet Durchsuchen
        foreach ($feld in $document.Fields){
            $text = $feld.code.Text
            $count += My_count_occurrence($text)
        }

        #Header Durchsuchen
        foreach ($header in $document.Sections[1].Headers){
            foreach ($feld in $header.Range.Fields){
                $text = $feld.code.Text
                $count += My_count_occurrence($text)
            }
        }

        #Footer Durchsuchen
        foreach ($footer in $document.Sections[1].Footers){
            foreach ($feld in $footer.Range.Fields){
                $text = $feld.code.Text
                $count += My_count_occurrence($text)
            }
        }
        if ($count -eq 0) #schliesen ohne speichern und Funktion beenden
        {
            try{
                #$document.Saved()
                $document.close($False) #schliesen ohne speichenr
            } catch {
                $application.quit()
                Write-Host "$path fehlerbehandung beim speichern" -ForegroundColor DarkMagenta
                return @($false, $path, "konnte nicht speichern", [System.ConsoleColor]::Red)
            }
        
            $application.quit()
            return @($false, $path, "anzhal felder $count", [System.ConsoleColor]::DarkGreen)
         }


        #---- wenn felder gefunden landen wir hier und ersetzten die felder

        #Dokumnet Felder Ersetzten
        foreach ($feld in $document.Fields){
            $text = $feld.code.Text
            $feld.code.Text = myReplace($text)
        }

        #Header Felder Ersetzten
        foreach ($header in $document.Sections[1].Headers){
            foreach ($feld in $header.Range.Fields){
                $text = $feld.code.Text
                $feld.code.Text = myReplace($text)
            }
        }

        #Footer Felder Ersetzten
        foreach ($footer in $document.Sections[1].Footers){
            foreach ($feld in $footer.Range.Fields){
                $text = $feld.code.Text
                $feld.code.Text = myReplace($text)
            }
        }

        try{
            #$document.Saved()
            $document.close($True) #schliesen und speichenr
        } catch {
            $application.quit()
            Write-Host "$path fehlerbehandung beim speichern" -ForegroundColor DarkMagenta
            return @($false, $path, "konnte nicht speichern", [System.ConsoleColor]::Red)
        }
        
        $application.quit()
        return @($false, $path, "anzhal felder gew. $count", [System.ConsoleColor]::DarkGreen)

    }
}



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


function myJobHandler()
{
    #fertige jobs
    $doneJob = Get-Job | Where-Object { $_.State -eq 'Completed' } | Select-Object -First 1
    if ($doneJob) {
        $erg = Receive-Job -Job $doneJob
        tocsv $erg[1] $erg[2] $erg[3]
        Remove-Job -Job $doneJob
        $global:jobs.Remove($doneJob.Id)
        return
    } 
    # überfällige jops
    $time_now = Get-Date
    $overcooked = $($global:jobs.GetEnumerator() | Sort-Object {$_.Value.time} | Select-Object -First 1 )
    $dTime = $($time_now - $overcooked.Value.time).TotalSeconds

    if ($dTime -ge $maxRunTime){
        #porzess killen
        $caption = [System.IO.Path]::GetFileNameWithoutExtension($overcooked.value.index)
        $prozess = Get-Process | Where-Object {($_.MainWindowTitle -like "*$("$caption")*") -and ($_.Name -eq "WINWORD")}
        if ($prozess){
            Stop-Process -Id $prozess.Id
        }
        Start-Sleep -Seconds 1
        Stop-Job -Id $overcooked.Name #Name ist Id :)
        Remove-Job -Id $overcooked.Name
        tocsv "Job $($overcooked.value.index)" "gestopt nach overtime $dTime s" -color red
        $global:jobs.Remove($overcooked.Name)

        return
    }

    Start-Sleep -Seconds 1
    Write-Host "."

}

#------------------------- Start Jops -----------------------
$initScript = $env
$wert = "Job"

$maxJobs = 2
$maxRunTime = 25
$maxLebenzeitWord = 30
$global:jobs = @{}

$output = ".\output.csv"
$path_file = ".\dir felder.txt"

$s=Get-Date


Start-Transcript -Path .\log.txt -Append
Get-Job | Remove-Job
$word_id = 24163
foreach($word_path in Get-Content $path_file -Encoding UTF8){
    $sd=Get-Date
    while ($global:jobs.Count -ge $maxJobs) {
        myJobHandler
    }


    $script = 
    {
        param($para1)
        process_Word($para1)
    }

    #Start-Job
    Write-Host "Starte $word_path"
    $jop = Start-Job -InitializationScript $initScript -ScriptBlock $script -ArgumentList ($word_path)
  
    $Job_data = [PSCustomObject]@{
                time= Get-Date
                index=$word_path
                word_id=$word_id
                }
    $global:jobs.Add( $jop.Id, $Job_data)


    #PWord Porzesse killen welche durchrutschen
    $prozess = (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) | Sort-Object {$_.StartTime} | Select-Object -First 1
    if($prozess){
        If (((Get-Date) - ($prozess.StartTime)).TotalSeconds -ge $maxLebenzeitWord){
            Write-Host("Prozess {0} nach overtime beendet" -f $prozess.MainWindowTitle) -ForegroundColor Red
            $prozess.Kill()
    }
}


    $e=Get-Date
    $Dtime = ($e - $sd).TotalSeconds
    Write-Host "anzahl: $word_id durchlaufzeit $Dtime s"
    $word_id += 1
}

while ($global:jobs.Count -gt 0) {
    myJobHandler
}


$e=Get-Date
$time = ($e - $s).TotalSeconds
Write-Host "runntime $time"
Stop-Transcript