
$csv = @()
$File = "MDB_STAMMDATEN.XML"
$ExportFile = "Daten.CSV"
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

# Load the XML content from the file
[xml]$xml = Get-Content (Join-Path -Path $ScriptDir -ChildPath $File)

# Define the XPath query to select only members of the 20th legislative period
$xpath = "//MDB[WAHLPERIODEN/WAHLPERIODE/WP='20']"

# Select the nodes that match the XPath query
$nodes = $xml.SelectNodes($xpath) #| Select-Object -First 100

Write-Output "Found $($nodes.Count) nodes"

foreach ($node in $nodes) {

    $name = $node.NAMEN.NAME

    if ( ($name | Measure-Object).Count -gt 1) {
        $nachname = $name.NACHNAME[(($name.NACHNAME | Measure-Object).Count - 1)]
        $vorname = $name.VORNAME[(($name.VORNAME | Measure-Object).Count - 1)]
    }
    else {

        $nachname = $node.NAMEN.NAME.NACHNAME
        $vorname = $node.NAMEN.NAME.VORNAME
    }
    $anredeTitel = [string]$node.NAMEN.NAME.ANREDE_TITEL
    $akadTitel = [string]$node.NAMEN.NAME.AKAD_TITEL
    $parteiKurz = $node.BIOGRAFISCHE_ANGABEN.PARTEI_KURZ
    $geschlecht = $node.BIOGRAFISCHE_ANGABEN.GESCHLECHT
    $id = $node.ID


    $childNode = $node.SelectSingleNode("WAHLPERIODEN/WAHLPERIODE[WP='20']")
    $wahlperiode = $childNode.WP
    $wahlperiodeBIS = $childNode.MDBWP_BIS

    switch ($geschlecht) {
        "männlich" {
            $anredeBrief = "Sehr geehrter"

            if ($anredeTitel -eq '') {
                $anredeBriefLang = ('Sehr geehrter Herr {0}' -f $nachname )
                $anschrift = ('{0} {1}' -f $vorname, $nachname )
            }
            else {
                $anredeBriefLang = ('Sehr geehrter Herr {0} {1}' -f $anredeTitel, $nachname )
                $anschrift = ('{0} {1} {2}' -f $anredeTitel, $vorname, $nachname )
            }
        }
        "weiblich" {
            $anredeBrief = "Sehr geehrte"
            if ($anredeTitel -eq '') {
                $anredeBriefLang = ('Sehr geehrte Frau {0}' -f $nachname )
                $anschrift = ('{0} {1}' -f $vorname, $nachname )
            }
            else {
                $anredeBriefLang = ('Sehr geehrte Frau {0} {1}' -f $anredeTitel, $nachname )
                $anschrift = ('{0} {1} {2}' -f $anredeTitel, $vorname, $nachname )
            }
        }
        default { $anredeBrief = "Sehr geehrte/r" }
    }

    if ($wahlperiodeBIS -eq '') {
        $property = [ordered]@{
            "ID"                = $id
            "Nachname"          = $nachname.Trim()
            "Vorname"           = $vorname.Trim()
            "AnredeTitel"       = $anredeTitel.Trim()
            "AkademischerTitel" = $akadTitel.Trim()
            "ParteiKurz"        = $parteiKurz
            "Geschlecht"        = $geschlecht
            "Wahlperiode"       = $wahlperiodeBIS
            "Anschrift"         = $anschrift
            "AnredeBrief"       = $anredeBrief
            "AnredeBriefLang"   = $anredeBriefLang
            "Gedruckt"          = 'nein'
        }

        $csv += New-Object PSObject -Property $property
    }
}

$csv | Export-Csv -Path (Join-Path -Path $ScriptDir -ChildPath $ExportFile) -NoTypeInformation -Force -Confirm:$false -Encoding UTF8