
#***********
#* Enums
#***********

enum enmPartOfSpeech 
{
    NA
    NOUN
    VERB
    ADJECTIVE
    ADVERB
}

enum enmGramGender
{
    NA
    MASCULINE
    FEMININE
    NEUTRAL
}

enum enmGramNb
{
    NA
    SINGULAR
    PLURAL
    UNCOUNTABLE
}

enum enmTermType
{
    NA
    FULL_FORM
    SHORT_FORM
    ACRONYM
    ABBREVIATION
    PHRASE
    VARIANT
}

#**************
#* Classes
#**************



<# <tig>
<term>19 geheiligte Glyphen von Unau</term>
<termNote type="termId">rKXFjtJmMPW2ro4QJaw1Q0z04</termNote>
<note>Note</note>
<termNote type="partOfSpeech">NOUN</termNote>
<termNote type="grammaticalGender">FEMININE</termNote>
<termNote type="grammaticalNumber">PLURAL</termNote>
<termNote type="forbidden">false</termNote>
<termNote type="preferred">false</termNote>
<termNote type="exactMatch">true</termNote>
<termNote type="status">Approved</termNote>
<termNote type="caseSensitive">true</termNote>
<termNote type="createdBy">vendolis</termNote>
<termNote type="createdAt">1573580740334</termNote>
<termNote type="lastModifiedBy">vendolis</termNote>
<termNote type="lastModifiedAt">1580239290287</termNote>
<termNote type="shortTranslation">Glyphen von Unau</termNote>
<termNote type="termType">PHRASE</termNote>
</tig>
 #>

class TigEntry
{
    [string]$term
    [string]$shortTranslation
    [string]$note
    [string]$termId
    [string]$usageNote
    [enmPartOfSpeech]$partOfSpeech
    [enmGramGender]$grammaticalGender
    [enmGramNb]$grammaticalNumber
    [bool]$forbidden
    [bool]$preferred
    [bool]$exactMatch
    [string]$status
    [bool]$caseSensitive
    [enmTermType]$termType
}

<#
<termEntry id="d43cdfea-a965-4f3a-8a2f-b7aec8776c9e">
<descrip type="conceptId">ev3YD2MA0cV5YB80Yq1oyNQj3</descrip>
<descrip type="conceptDomain">TDE</descrip>
<descrip type="conceptNote">Rule term</descrip>
<descrip type="conceptSubdomain">General</descrip>
<descrip type="conceptUrl">Core Rules</descrip>
<langSet xml:lang="de"><tig></tig></langSet>
<langSet xml:lang="en-us"><tig></tig></langSet>
</termEntry>
#>

class TermEntry
{
    [string]$conceptId
    [string]$conceptDomain
    [string]$conceptNote
    [string]$conceptSubdomain
    [string]$conceptUrl
    $lang_de = [System.Collections.ArrayList]@()
    $lang_en = [System.Collections.ArrayList]@()
}


##################
#  Functions
##################

function New-TermID
{
    return $(([char[]]([char]65..[char]90) + ([char[]]([char]97..[char]122)) + 0..9 | sort {Get-Random})[0..24] -join '')
}


function New-TigListFromEntry {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        [string]
        $entry
    )
    
    $tigList = [System.Collections.ArrayList]@()

    [TigEntry]$mainTig = [TigEntry]::new()

    [string[]] $mainTags = (Select-String '\[[^\]]+\]' -input $entry -AllMatches).matches.Value

    if($mainTags)
    {
        $mainTig.term = ($entry.Split('[').Trim())[0]
        foreach ($tag in $mainTags)
        {
            switch -regex ($tag) 
            {
                '\[m.\]' { $mainTig.grammaticalGender = 'MASCULINE' }
                '\[f.\]' { $mainTig.grammaticalGender = 'FEMININE' }
                '\[n.\]' { $mainTig.grammaticalGender = 'NEUTRAL' }
                '\[acro.([^\]]+)\]' 
                { 
                    [TigEntry]$acroTig = [TigEntry]::new()
                    $acroTig.term = $($matches[1].Trim())
                    $acroTig.termType = 'ACRONYM'
                    $acroTig.caseSensitive = $true
                    $acroTig.exactMatch = $true
                    [void] $tigList.Add($acroTig)
                }
                '\[abbr.([^\]]+)\]' 
                { 
                    [TigEntry]$abbrTig = [TigEntry]::new()
                    $abbrTig.term = $($matches[1].Trim())
                    $abbrTig.termType = 'ABBREVIATION'
                    if($abbrTig.term -cmatch '[A-Z]')
                    {
                        $abbrTig.caseSensitive = $true
                    }
                    $abbrTig.exactMatch = $true
                    [void] $tigList.Add($abbrTig)
                }
                '\[note:([^\]]+)\]' { $mainTig.note = $($matches[1].Trim()) }
                '\[dep.\]' { $mainTig.forbidden = $true }
                '\[pl.([^\]]+)\]'
                {
                    $mainTig.grammaticalNumber = 'SINGULAR'
                    [TigEntry]$plTig = [TigEntry]::new()
                    $plTig.term = $($matches[1].Trim())
                    $plTig.grammaticalNumber = 'PLURAL'
                    if($plTig.term -cmatch '[A-Z]')
                    {
                        $plTig.caseSensitive = $true
                    }
                    [void] $tigList.Add($plTig)
                }
                '\[derog.\]' { $mainTig.usageNote = "Derogative" }
                '\[old\]' { $mainTig.forbidden = $true }
                '\[adj.([^\]]+)\]' 
                { 
                    $mainTig.partOfSpeech = 'NOUN'
                    [TigEntry]$adjTig = [TigEntry]::new()
                    $adjTig.term = $($matches[1].Trim())
                    $adjTig.partOfSpeech = 'ADJECTIVE'
                    if($adjTig.term -cmatch '[A-Z]')
                    {
                        $adjTig.caseSensitive = $true
                    }
                    [void] $tigList.Add($adjTig)
                }
                '\[pref.\]' { $mainTig.preferred = $true }
                '\[alt.([^\]]*)\]'
                { 
                    if(!$($matches[1].Trim()))
                    {
                        $mainTig.preferred = $false
                        $mainTig.termType = 'VARIANT'
                    }
                    else 
                    {
                        $mainTig.preferred = $true
                        [TigEntry]$altTig = [TigEntry]::new()
                        $altTig.term = $($matches[1].Trim())
                        if($altTig.term -cmatch '[A-Z]')
                        {
                            $altTig.caseSensitive = $true
                        }
                        $altTig.termType = 'VARIANT'
                        [void] $tigList.Add($altTig)    
                    }
                }
                Default { "$_ is not a known tag." }
            }
        }
    }
    else 
    {
        $mainTig.term = $entry.Trim()
    }

    if($mainTig.term -cmatch '[A-Z]')
    {
        $mainTig.caseSensitive = $true
    }

    [void] $tigList.Add($mainTig)

    return $tigList

}

function New-EntryFromMasterTable {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true)]
        $entry
    )
    
    [TermEntry] $newEntry = [TermEntry]::new()

    $newEntry.conceptDomain = "TDE"
    $newEntry.conceptNote = $entry.Notes
    $newEntry.conceptUrl = $entry.Location

    # Handling the German Side
    [string[]] $germanEntries = $entry.German.Split(';').Trim()

    foreach ($gEntry in $germanEntries | Where-Object {$_}) 
    {
        $gTigList = New-TigListFromEntry($gEntry)

        foreach ($tig in $gTigList)
        {
            [void] $newEntry.lang_de.Add($tig)
        }
    }

    # Handle English entries

    [string[]] $engMainEntries = $entry.'Approved English Term'.Split(';').Trim()

    if($entry.'Alternative Translations')
    {
        [string[]] $engAltEntries = $entry.'Alternative Translations'.Split(';').Trim()
        $hasAltEntries = $true
    }
    else 
    {
        $engAltEntries = $null
        $hasAltEntries = $false
    }

 
    foreach ($engEntry in $engMainEntries | Where-Object {$_})
    {
        if($hasAltEntries)
        {
            $engEntry += " [pref.]"
        }

        $engTigList = New-TigListFromEntry($engEntry)

        foreach ($tig in $engTigList)
        {
            [void] $newEntry.lang_en.Add($tig)
        }
    }

    foreach ($engEntry in $engAltEntries | Where-Object {$_})
    {
        $engEntry += " [alt.]"

        $engTigList = New-TigListFromEntry($engEntry)

        foreach ($tig in $engTigList)
        {
            [void] $newEntry.lang_en.Add($tig)
        }
    }

    return $newEntry
}

function ConvertTo-TBX {
    param (
        $TermList
    )

    $fullXML = ""
    $xmlHeader = @"
<?xml version="1.0" encoding="UTF-8"?>
<martif xml:lang="en" type="TBX">
<text>
<body>
"@

    $xmlFooter = @"
</body>
</text>
</martif>
"@

    $fullXML += $xmlHeader

    foreach($entry in $TermList)
    {
        $fullXML += "`r`n" + $(New-XMLTermEnty($entry))
    }
    
    $fullXML += "`r`n" + $xmlFooter

    return $fullXML
}

function New-XMLTermEnty {
    param (
        [TermEntry]
        $termEntry
    )
    
    [string] $lang_de = ""
    [string] $lang_en = ""
    
    foreach($deTig in $termEntry.lang_de)
    {
        if($lang_de)
        {
            $lang_de += "`r`n"
        }

        $lang_de += $(New-XMLTigEntry($deTig))
    }

    foreach($enTig in $termEntry.lang_en)
    {
        if($lang_en)
        {
            $lang_en += "`r`n"
        }

        $lang_en += $(New-XMLTigEntry($enTig))
    }

    $xmlEntry = @"
<termEntry id="$((New-Guid).Guid)">
<descrip type="conceptId">$(New-TermID)</descrip>
<descrip type="conceptDomain">$([Security.SecurityElement]::Escape($termEntry.conceptDomain))</descrip>
<descrip type="conceptNote">$([Security.SecurityElement]::Escape($termEntry.conceptNote))</descrip>
<descrip type="conceptSubdomain">$([Security.SecurityElement]::Escape($termEntry.conceptSubdomain))</descrip>
<descrip type="conceptUrl">$([Security.SecurityElement]::Escape($termEntry.conceptUrl))</descrip>
<langSet xml:lang="de">
$($lang_de)
</langSet>
<langSet xml:lang="en">
$($lang_en)
</langSet>
</termEntry>
"@

    return $xmlEntry
}

function New-XMLTigEntry {
    param (
        [TigEntry]
        $tigEntry
    )
    

    $xmlEntry = @"
<tig>
<term>$([Security.SecurityElement]::Escape($tigEntry.term))</term>
<termNote type="termId">$(New-TermID)</termNote>
$(if($tigEntry.termId){'<termNote type="termId">' + $tigEntry.termId + '</termNote>'})
<note>$([Security.SecurityElement]::Escape($tigEntry.note))</note>
$(if($tigEntry.partOfSpeech -ne 'NA'){'<termNote type="partOfSpeech">' + $tigEntry.partOfSpeech + '</termNote>'})
$(if($tigEntry.grammaticalGender -ne 'NA'){'<termNote type="grammaticalGender">' + $tigEntry.grammaticalGender + '</termNote>'})
$(if($tigEntry.grammaticalNumber -ne 'NA'){'<termNote type="grammaticalNumber">' + $tigEntry.grammaticalNumber + '</termNote>'})
<termNote type="forbidden">$($tigEntry.forbidden)</termNote>
$(if($tigEntry.preferred){'<termNote type="preferred">' + $tigEntry.preferred + '</termNote>'})
<termNote type="exactMatch">$($tigEntry.exactMatch)</termNote>
$(if($tigEntry.caseSensitive){'<termNote type="caseSensitive">' + $tigEntry.caseSensitive + '</termNote>'})
$(if($tigEntry.shortTranslation){'<termNote type="shortTranslation">' + $([Security.SecurityElement]::Escape($tigEntry.shortTranslation)) + '</termNote>'})
$(if($tigEntry.termType -ne 'NA'){'<termNote type="termType">' + $tigEntry.termType + '</termNote>'})
</tig>
"@ -replace "[`r`n]+`r`n", "`r`n"

    return $xmlEntry
}

#*********************
#* Skript 
#*********************

Get-ExcelSheetInfo C:\Users\Tajo\OneDrive\Ulisses_Work\MasterGlossary.xlsx

$MasterTable = Import-Excel C:\Users\Tajo\OneDrive\Ulisses_Work\MasterGlossary.xlsx -WorksheetName MasterTable
$KingdomsCitiesPeople = Import-Excel C:\Users\Tajo\OneDrive\Ulisses_Work\MasterGlossary.xlsx -WorksheetName Kingdoms-Cities-People
$Abbreviations = Import-Excel C:\Users\Tajo\OneDrive\Ulisses_Work\MasterGlossary.xlsx -WorksheetName Abbreviations

######################
# Process MasterTable
######################

$MasterList = [System.Collections.ArrayList]@()

foreach ($entry in $MasterTable)
{
    [void] $MasterList.Add($(New-EntryFromMasterTable($entry)))
    Write-Information "Loaded $($MasterList.Count) entries from master table."
}

$fullXML = ConvertTo-TBX($MasterList)

Set-Content glossary.tbx -Value $fullXML

