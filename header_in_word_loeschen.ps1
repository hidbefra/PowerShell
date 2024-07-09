Set-StrictMode -Version latest

$path_file = ".\path.txt"
$output = ".\output.csv"

Function prof8_header_in_word_loeschen
{

    Foreach($file in Get-Content $path_file -Encoding UTF8)
    {
        my_prozess $file
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


Function my_prozess ($file){
    
     Write-Host $file

    try {
        [System.Collections.ArrayList]$doc = [System.IO.File]::ReadAllBytes($file)
        }
    catch {
        tocsv $file "kann nicht geöffnet werden" -color red
        continue
    }

    #prüfen, ob datei mit FSServer startet
    $start_string = [System.Text.Encoding]::UTF8.GetString($doc[0..7])
    if ( $start_string -ne "FSServer"){
        tocsv $file "String FSServer nicht gefunden" -color red
        continue
    }

    #die ersten Byts löschen
    $index = $doc.IndexOf([convert]::ToByte(208))
    $doc.RemoveRange(0,$index)

    #prüfen ob datei mit dem richtigen wert startet
    $byt1 = $doc[0..3]
    $byt2 = [byte[]]@(0xd0, 0xcf, 0x11, 0xe0)
    $str1 = [System.Text.Encoding]::UTF8.GetString($byt1)
    $str2 = [System.Text.Encoding]::UTF8.GetString($byt2)
    if ($str1 -ne $str2){
        tocsv $file "nach löschen ware das erste zeichen nicht korrekt" -color red
        continue
    }

    #zurückspeichern
    try {
        Set-Content -Path $file -Value $doc -Encoding Byte
         }
    catch {
        tocsv $file "kann nicht gespeichert werden" -color red
        continue
    }
    tocsv $file "modifiziert" -color DarkGreen

}


Start-Transcript -Path .\log.txt -Append
prof8_header_in_word_loeschen
Stop-Transcript