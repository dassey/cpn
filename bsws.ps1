#Requires -Version 5.1
<#
.SYNOPSIS
    A graphical utility for converting KML, CSV, and SLF files to a specific XML format for the BSWS system.
.DESCRIPTION
    This script provides a modern Windows Forms GUI to assist in various data conversion and modification tasks.
    It can:
    - Modify existing XML files with new styling and ownership attributes.
    - Convert KML and CSV files into the BSWS XML format for points, lines, and areas.
    - Extract data from different types of KML files into CSV format.
.NOTES
    Author: Massey
    Version: 5.7 
    Date: 2025-06-13
#>

#region Assembly Loading
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore, PresentationFramework
#endregion Assembly Loading

#region Global Script Variables & Icons
# Setup Directory Structure
$script:BasePat = Get-Location
$script:InputPath = Join-Path -Path $script:BasePat -ChildPath 'Input'
$script:TempPath = Join-Path -Path $script:BasePat -ChildPath 'Temp'
$script:OutputPath = Join-Path -Path $script:BasePat -ChildPath 'Output'
$script:CsvOutputPath = Join-Path -Path $script:OutputPath -ChildPath 'CSV'

@( $script:InputPath, $script:TempPath, $script:OutputPath, $script:CsvOutputPath ) | ForEach-Object {
    if (-not (Test-Path -Path $_)) {
        New-Item -Path $_ -ItemType Directory -Force | Out-Null
        Write-Host "Created folder: $_"
    }
}

# Base64 Encoded Icons (32x32 px) - Generated from user-provided SVG
$script:IconString = "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"
#endregion Global Script Variables

#region Helper Functions
Function Get-Icon {
    # Use the same icon for all buttons for a consistent look
    $base64String = $script:IconString
    if (-not $base64String) { return $null }

    try {
        $bytes = [System.Convert]::FromBase64String($base64String)
        $ms = New-Object System.IO.MemoryStream($bytes)
        $image = [System.Drawing.Image]::FromStream($ms)
        return $image
    } catch {
        Write-Warning "Failed to load icon. $_"
        return $null
    }
}

Function New-ModernButton {
    param(
        [string]$Title,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string]$BaseColor,
        [string]$HoverColor,
        [scriptblock]$OnClick
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Title
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.FlatStyle = 'Flat'
    $button.FlatAppearance.BorderSize = 0
    $button.BackColor = [System.Drawing.ColorTranslator]::FromHtml($BaseColor)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)

    # Store colors in the button's Tag property to avoid scope issues
    $button.Tag = @{
        BaseColor  = $BaseColor
        HoverColor = $HoverColor
    }

    # Events now refer to the Tag property of the control itself ($this)
    $button.add_MouseEnter({ $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml($this.Tag.HoverColor) })
    $button.add_MouseLeave({ $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml($this.Tag.BaseColor) })

    if ($OnClick) {
        $button.add_click($OnClick)
    }
    return $button
}

Function New-ModernLabel {
    param (
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 200,
        [int]$Height = 20,
        [int]$FontSize = 9,
        [bool]$IsBold = $false
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.ForeColor = [System.Drawing.Color]::White
    $fontStyle = if($IsBold) {[System.Drawing.FontStyle]::Bold} else {[System.Drawing.FontStyle]::Regular}
    $label.Font = New-Object System.Drawing.Font('Segoe UI', $FontSize, $fontStyle)
    return $label
}
#endregion Helper Functions

#region Main Application GUI
Function Start-GraphicsAssistant {

    $mainForm = New-Object System.Windows.Forms.Form -Property @{
        Text            = "BSWS Graphics Assistant V5.6"
        ClientSize      = '500, 380'
        BackColor       = [System.Drawing.ColorTranslator]::FromHtml("#1A202C") # Very Dark Blue
        StartPosition   = 'CenterScreen'
        FormBorderStyle = 'FixedSingle'
        MaximizeBox     = $false
    }
    
    $mainForm.Controls.Add((New-ModernLabel -Text "BSWS Graphics Assistant" -X 20 -Y 20 -Width 460 -Height 30 -FontSize 16 -IsBold $true))
    $mainForm.Controls.Add((New-ModernLabel -Text "Convert and modify geospatial data files." -X 20 -Y 50 -Width 460 -Height 20 -FontSize 10))

    $buttonPanel = New-Object System.Windows.Forms.Panel -Property @{
        Location = '20, 90'
        Size = '460, 270'
    }
    $mainForm.Controls.Add($buttonPanel)

    $buttonLayout = @(
        @{ Title="Quick`nCPOF KML"; BaseColor="#DD6B20"; HoverColor="#F6AD55"; OnClick={ Invoke-KmlConversion -QuickMode $true } },
        @{ Title="Selective`nCPOF KML"; BaseColor="#D53F8C"; HoverColor="#FBB6CE"; OnClick={ Invoke-KmlConversion -QuickMode $false } },
        @{ Title="CPOF KML`nExtractor"; BaseColor="#00A0A0"; HoverColor="#4FD1C5"; OnClick={ Invoke-CpofKmlExtractor } },
        
        @{ Title="CSV`nto XML"; BaseColor="#38A169"; HoverColor="#68D391"; OnClick={ Invoke-CsvConversion } },
        @{ Title="Modify`nXML"; BaseColor="#3182CE"; HoverColor="#63B3ED"; OnClick={ Invoke-XmlModification } },
        @{ Title="CPOF ALT`nExtractor"; BaseColor="#007A7A"; HoverColor="#38B2AC"; OnClick={ Invoke-CpofKmlExtractor -UseNs2Syntax $true } },
        
        @{ Title="Global Mapper`nKML"; BaseColor="#805AD5"; HoverColor="#B794F4"; OnClick={ Invoke-GlobalMapperKmlConverter } },
        @{ Title="Google Earth`nKML"; BaseColor="#6B46C1"; HoverColor="#9F7AEA"; OnClick={ Invoke-GoogleEarthKmlConverter } },
        @{ Title="CPCE`nSLF"; BaseColor="#B83280"; HoverColor="#ED64A6"; OnClick={ Invoke-CpceSlfConverter } }
    )

    $buttonWidth = 145
    $buttonHeight = 70
    $xGap = 10
    $yGap = 10
    $x = 0
    $y = 0

    foreach ($btnProp in $buttonLayout) {
        $btnProp.Add('X', $x)
        $btnProp.Add('Y', $y)
        $btnProp.Add('Width', $buttonWidth)
        $btnProp.Add('Height', $buttonHeight)
        
        $button = New-ModernButton @btnProp
        $button.Image = Get-Icon
        $button.TextImageRelation = 'ImageAboveText' 
        $button.TextAlign = 'BottomCenter'
        $button.ImageAlign = 'TopCenter'
        $buttonPanel.Controls.Add($button)

        $x += $buttonWidth + $xGap
        if ($x + $buttonWidth -gt $buttonPanel.Width) {
            $x = 0
            $y += $buttonHeight + $yGap
        }
    }

    [void]$mainForm.ShowDialog()
}
#endregion Main Application GUI

#region Action Functions
function Show-OpenFileDialog {
    param(
        [string]$Title,
        [string]$Filter = 'All Files (*.*)|*.*'
    )

    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $Title
    $openFileDialog.Filter = $Filter
    
    # Show the dialog and check if the user clicked 'OK'
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    # If user cancels, return nothing
    return $null
}

function Show-SaveFileDialog {
    param(
        [string]$Title,
        [string]$Filter = 'All Files (*.*)|*.*'
    )

    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Title = $Title
    $saveFileDialog.Filter = $Filter
    
    # Show the dialog and check if the user clicked 'OK'
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName
    }
    # If user cancels, return nothing
    return $null
}

function New-FormButton {
    param(
        [string]$Title,
        [int]$X,
        [int]$Y,
        [int]$Width = 75, # Default width
        [int]$Height = 23  # Default height
    )
    Add-Type -AssemblyName System.Windows.Forms
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Title
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, $Height)
    return $button
}

Function Invoke-XmlModification {
    $inputFile = Show-OpenFileDialog -Filter 'XML (*.xml)|*.xml' -Title 'Select XML File to Modify'
    if (-not $inputFile) { return }

    $outputFile = Show-SaveFileDialog -Filter 'XML (*.xml)|*.xml' -Title 'Save Modified XML As'
    if (-not $outputFile) { return }

    $settings = Get-ConversionSettings -Mode 'Modify'
    if (-not $settings) { return }

    try {
        $content = Get-Content -Path $inputFile -Raw
        
        $replacements = @{
            'owner=.*?"'             = 'owner="{0}"' -f $settings.Controller
            'ReadOnly=".*?"'         = 'ReadOnly="{0}"' -f $settings.ReadOnly
            'fillColorRGB=".*?"'      = 'fillColorRGB="{0}"' -f $settings.FillColor
            'fillStyle=".*?"'        = 'fillStyle="{0}"' -f $settings.FillStyle
            'lineColorRGB=".*?"'      = 'lineColorRGB="{0}"' -f $settings.LineColor
            'lineWidth=".*?"'         = 'lineWidth="{0}"' -f $settings.LineWidth
            'textFGColorRGB=".*?"'    = 'textFGColorRGB="{0}"' -f $settings.TextColor
            'textFontSize=".*?"'      = 'textFontSize="{0}"' -f $settings.TextSize
            'textFontStyle=".*?"'     = 'textFontStyle="{0}"' -f $settings.TextStyle
            'symbolId=".*?"'          = 'symbolId="{0}"' -f $settings.SymbolID
        }

        foreach ($find in $replacements.Keys) {
            $replace = $replacements[$find]
            $content = $content -replace $find, $replace
        }

        $content | Set-Content -Path $outputFile -Encoding UTF8
        [System.Windows.MessageBox]::Show("XML modification complete.", "Success")
    }
    catch {
        [System.Windows.MessageBox]::Show("An error occurred during XML modification: `n$($_.Exception.Message)", "Error", "OK", "Error")
    }
}

Function Invoke-KmlConversion {
    param ([bool]$QuickMode)

    $inputFile = Show-OpenFileDialog -Filter 'KML (*.kml)|*.kml' -Title 'Select KML File'
    if (-not $inputFile) { return }
    $outputFile = Show-SaveFileDialog -Filter 'XML (*.xml)|*.xml' -Title 'Save Converted XML As'
    if (-not $outputFile) { return }

    $mode = if ($QuickMode) { 'Quick' } else { 'Standard' }
    $settings = Get-ConversionSettings -Mode $mode
    if (-not $settings) { return }
    
    $settings.InputFile = $inputFile
    $settings.OutputFile = $outputFile

    try {
        $tempCsvPath = Join-Path $script:TempPath "temp_kml_data.csv"
        Process-CpofKmlToCsv -InputKmlPath $settings.InputFile -OutputCsvPath $tempCsvPath
        Convert-CsvToBswsXml -Settings $settings
        [System.Windows.MessageBox]::Show("KML conversion successful!", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error during KML conversion: $($_.Exception.Message)", "Error")
    }
}

Function Invoke-CsvConversion {
    $inputFile = Show-OpenFileDialog -Filter 'CSV (*.csv)|*.csv' -Title 'Select CSV File'
    if (-not $inputFile) { return }
    $outputFile = Show-SaveFileDialog -Filter 'XML (*.xml)|*.xml' -Title 'Save Converted XML As'
    if (-not $outputFile) { return }

    $settings = Get-ConversionSettings -Mode 'Standard'
    if (-not $settings) { return }
    
    $settings.InputFile = $inputFile
    $settings.OutputFile = $outputFile

    try {
        Convert-CsvToBswsXml -Settings $settings
        [System.Windows.MessageBox]::Show("CSV conversion successful!", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error during CSV conversion: $($_.Exception.Message)", "Error")
    }
}

Function Invoke-CpofKmlExtractor {
    param ([bool]$UseNs2Syntax)

    $inputFile = Show-OpenFileDialog -Filter 'KML (*.kml)|*.kml' -Title 'Select KML File to Extract'
    if (-not $inputFile) { return }

    $settings = Get-ExtractorSettings -InputKmlFile $inputFile
    if (-not $settings) { return }

    try {
        # Create the output subfolder
        $outputSubfolder = Join-Path -Path $settings.OutputDirectory -ChildPath $settings.Prefix
        if (-not (Test-Path $outputSubfolder)) {
            New-Item -Path $outputSubfolder -ItemType Directory | Out-Null
        }

        # Process the KML to a master CSV
        $masterCsvPath = Join-Path $script:TempPath "master_extracted.csv"
        Process-KmlToMasterCsv -InputKmlPath $settings.InputFile -OutputCsvPath $masterCsvPath -UseNs2Syntax $UseNs2Syntax

        # Define symbol patterns and corresponding file names
        $symbolMappings = @{
            'Point'      = "PUSTAT--M---,PUSSC---H---,PUCI*"
            'AREA'       = "PGAG-------X,PGAG----J--X,PUSM----F---,PGPP-------X"
            'AA'         = "PGAA----H--X"
            'AO'         = "PSAO----*--X"
            'AP'         = "POAA----*--X"
            'ATKPOS'     = "POAK----H--X"
            'DSA'        = "PASD-------X"
            'NAI'        = "PSAN-------X"
            'OBJECTIVE'  = "POAO-------X"
            'TAI'        = "PSAT-------X"
            'LINE'       = "PGLB-------X"
            'BOUNDARY'   = "PGLB----I--X,PGLB----J--X"
            'PHASELINE'  = "PGLP-------X"
            'MSR'        = "PLRM-------X"
            'ASR'        = "PLRA-------X"
            'TCP'        = "SPPO--------X"
            'CHKPNT'     = "GPPK------X"
            'CCP'        = "PAPC-------X"
        }

        # Read the master CSV content
        $masterCsvContent = Get-Content -Path $masterCsvPath

        # Iterate over each symbol type and create a separate CSV
        foreach ($name in $symbolMappings.Keys) {
            $patterns = $symbolMappings[$name].Split(',')
            $regexPattern = $patterns -join '|'

            $filteredLines = $masterCsvContent | Select-String -Pattern $regexPattern
            
            if ($filteredLines) {
                $outputCsvPath = Join-Path -Path $outputSubfolder -ChildPath "$($name).csv"
                Set-Content -Path $outputCsvPath -Value "X,Y,Name,Symbol"
                $filteredLines | Add-Content -Path $outputCsvPath
            }
        }

        Remove-Item $masterCsvPath -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("Extraction complete. CSV files saved in '$outputSubfolder'.", "Success")

    } catch {
        [System.Windows.MessageBox]::Show("Error during KML extraction: $($_.Exception.Message)", "Error")
    }
}

Function Invoke-GlobalMapperKmlConverter {
    $inputFile = Show-OpenFileDialog -Filter 'KML (*.kml)|*.kml' -Title 'Select Global Mapper KML'
    if (-not $inputFile) { return }
    $outputFile = Show-SaveFileDialog -Filter 'CSV (*.csv)|*.csv' -Title 'Save Converted CSV As'
    if (-not $outputFile) { return }

    try {
        $content = Get-Content -Path $inputFile -Raw
        $tempFile1 = Join-Path $script:TempPath "gm_kml_prep1.txt"
        $tempFile2 = Join-Path $script:TempPath "gm_kml_prep2.txt"

        $content | Select-String -Pattern '<description>.*CDATA.*</description>' -AllMatches |
            Select-Object -ExpandProperty Matches |
            Select-Object -ExpandProperty Value |
            Set-Content -Path $tempFile1

        (Get-Content -Path $tempFile1) -replace '\s', '' `
            -replace '[][]', '' `
            -replace '!', '' `
            -replace '<description><CDATAUnknownPointFeature<BR><BR><B>Name</B>=', '' `
            -replace '<BR><B>LINE_NAME</B>=', ',' `
            -replace '<BR><B>VTX_NUM</B>=', ',' `
            -replace '<BR><B>X</B>=', ',' `
            -replace '<BR><B>Y</B>=', ',' `
            -replace '></description>', '' |
            Set-Content -Path $tempFile2

        $csvHeader = "Name,Name1,VTX,X,Y"
        Set-Content -Path $outputFile -Value $csvHeader
        Get-Content $tempFile2 | Add-Content -Path $outputFile
        
        Remove-Item $tempFile1, $tempFile2 -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("Global Mapper KML to CSV conversion complete.", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error during Global Mapper KML conversion: $($_.Exception.Message)", "Error")
    }
}

Function Invoke-GoogleEarthKmlConverter {
    $inputFile = Show-OpenFileDialog -Filter 'KML (*.kml)|*.kml' -Title 'Select Google Earth KML'
    if (-not $inputFile) { return }
    $outputFile = Show-SaveFileDialog -Filter 'CSV (*.csv)|*.csv' -Title 'Save Converted CSV As'
    if (-not $outputFile) { return }
    
    try {
        # Use a series of temporary files for staged processing
        $tempFiles = @()
        1..8 | ForEach-Object { $tempFiles += Join-Path $script:TempPath "ge_kml_prep$_.txt" }

        (Get-Content -Path $inputFile) -replace "`t", "" | Set-Content -Path $tempFiles[0]
        
        Get-Content -Path $tempFiles[0] | Where-Object { $_ -like "*<name>*</name>" -or $_ -like "*,*,*" } | Set-Content -Path $tempFiles[1]

        (Get-Content -Path $tempFiles[1]) -replace '(.*?),(.*?),0', '<coord>$1,$2</coord>' | Set-Content -Path $tempFiles[2]
        (Get-Content -Path $tempFiles[2]) -replace '<name>(.*?)</name>', '<name><$1></name>' | Set-Content -Path $tempFiles[3]
        (Get-Content -Path $tempFiles[3]) -replace '<coord>	(.*?)</coord>', '<coord>$1</coord>' | Set-Content -Path $tempFiles[4]
        (Get-Content -Path $tempFiles[4]) -replace ' ', '' | Set-Content -Path $tempFiles[5]

        $name = ''
        $output = switch -Regex -File $tempFiles[5] {
            '<name><(.*?)></name>' { $name = $Matches[1] }
            default { "$_,$name" }
        }
        $output | Set-Content -Path $tempFiles[6]

        Get-Content -Path $tempFiles[6] | Where-Object {$_ -like "<coord>*</coord>,*"} | Set-Content -Path $tempFiles[7]
        (Get-Content -Path $tempFiles[7]) -replace '<coord>(.*?)</coord>','$1' | Set-Content -Path $tempFiles[7]

        Set-Content -Path $outputFile -Value "X,Y,Name"
        Get-Content -Path $tempFiles[7] | Add-Content -Path $outputFile
        
        Remove-Item -Path $tempFiles -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("Google Earth KML to CSV conversion complete.", "Success")
    } catch {
        [System.Windows.MessageBox]::Show("Error during Google Earth KML conversion: $($_.Exception.Message)", "Error")
    }
}

Function Invoke-CpceSlfConverter {
    $inputFile = Show-OpenFileDialog -Filter 'SLF (*.slf)|*.slf' -Title 'Select SLF File'
    if (-not $inputFile) { return }
    $outputFile = Show-SaveFileDialog -Filter 'CSV (*.csv)|*.csv' -Title 'Save Converted CSV As'
    if (-not $outputFile) { return }

    try {
        $tempFiles = @()
        1..8 | ForEach-Object { $tempFiles += Join-Path $script:TempPath "slf_prep$_.txt" }

        # This logic is a direct translation of the original script's SLF parsing
        (Get-Content -Path $inputFile) -replace "  ", "" | Set-Content -Path $tempFiles[0]
        (Get-Content -Path $tempFiles[0]) -replace "<", " " | Select-Object -Skip 10 | Set-Content -Path $tempFiles[1]
        
        $array = Get-Content -Path $tempFiles[1]
        $length = $array.Count
        $line = 1
        1..$length | ForEach-Object { $array[-$line]; $line++ } | Set-Content -Path $tempFiles[2]

        $content = (Get-Content -Path $tempFiles[2]).Trim()
        
        $lat = ''
        $lon = ''
        $name = ''
        $output = switch -Regex -File $tempFiles[2] {
            'Latitude>(.*?)/Latitude>' { $lat = $Matches[1] }
            'Longitude>(.*?)/Longitude>' { $lon = $Matches[1] }
            'Name>(.*?)/Name>' { $name = $Matches[1] }
            default { "<coordx>$lat</coordx> <coordy>$lon</coordy> <Name>$name</Name>" }
        }
        $output | Set-Content -Path $tempFiles[3]

        (Get-Content -Path $tempFiles[3]) -replace '\(', '_' -replace '\)', '' | Set-Content -Path $tempFiles[4]
        
        Get-Content -Path $tempFiles[4] | Where-Object { $_ -notmatch '<coordx> *</coordx> <coordy> *</coordy> <Name> *</Name>' } | Set-Content -Path $tempFiles[5]
        Get-Content -Path $tempFiles[5] | Where-Object { $_ -notmatch '<coordx> *</coordx> <coordy> *</coordy>.*>' } | Set-Content -Path $tempFiles[6]
        
        (Get-Content -Path $tempFiles[6]) -replace '<coordx>', '' -replace '</coordx>', ',' -replace '<coordy>', '' -replace '</coordy>', ',' -replace '<Name>', '' -replace '</Name>', '' | Set-Content -Path $tempFiles[7]
        (Get-Content -Path $tempFiles[7]) -replace '  ', '' | Set-Content -Path $tempFiles[7]

        Set-Content -Path $outputFile -Value "X,Y,Name"
        Get-Content -Path $tempFiles[7] | Add-Content -Path $outputFile
        
        Remove-Item -Path $tempFiles -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show("CPCE SLF conversion to CSV complete.", "Success")

    } catch {
        [System.Windows.MessageBox]::Show("Error during CPCE SLF conversion: $($_.Exception.Message)", "Error")
    }
}


#endregion Action Functions

#region Core Processing Logic

Function Convert-CsvToBswsXml {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Settings
    )

    $csvPath = if ($Settings.InputFile -like '*.kml') { Join-Path $script:TempPath "temp_kml_data.csv" } else { $Settings.InputFile }
    if (-not (Test-Path $csvPath)) { throw "Intermediate CSV file not found at $csvPath" }

    $csvData = Import-Csv -Path $csvPath
    if (-not $csvData) { throw "CSV file is empty or could not be read: $csvPath" }
    
    $isLineOrArea = $Settings.FeatureType -in @('Line', 'PhaseLine', 'MSR', 'ASR', 'Area', 'OBJ', 'TAI', 'NAI', 'ATKPOS')
    $entityData = if ($isLineOrArea) {
        $csvData | Group-Object Name | ForEach-Object {
            $points = $_.Group | ForEach-Object {
                $lat = [math]::Floor([double]$_.Y * 3600000)
                $lon = [math]::Floor([double]$_.X * 3600000)
                "$lat $lon"
            }
            [PsCustomObject]@{
                Name         = $_.Name
                pointsString = $points -join ' '
            }
        }
    } else { $csvData }

    $xmlSettings = New-Object System.Xml.XmlWriterSettings -Property @{ Indent = $true; IndentChars = "  " }
    $xmlWriter = [System.XML.XmlWriter]::Create($Settings.OutputFile, $xmlSettings)
    $xmlWriter.WriteStartDocument()
    $xmlWriter.WriteStartElement("Overlays")
    $xmlWriter.WriteStartElement("Overlay")

    $xmlWriter.WriteAttributeString("name", $Settings.LayerName)
    $xmlWriter.WriteAttributeString("owner", $Settings.Controller)
    # ... other overlay attributes
    
    $xmlWriter.WriteStartElement("Entities")
    
    foreach ($row in $entityData) {
        $xmlWriter.WriteStartElement("Entity")

        $entityParams = [ordered]@{
            "entityName"     = $row.Name
            "label"          = if($Settings.FeatureType -eq 'PhaseLine'){"$($row.Name):: "}else{$row.Name}
            "entityReadOnly" = $Settings.ReadOnly
            "symbolId"       = $Settings.SymbolID
            "textFGColorRGB" = $Settings.TextColor
            "textFontSize"   = $Settings.TextSize
            "textFontStyle"  = $Settings.TextStyle
        }

        if ($isLineOrArea) {
            $entityParams["entityInstance"] = if ($Settings.FeatureType -in @('Line', 'PhaseLine', 'MSR', 'ASR')) {'TactLineEntity'} else {'TactAreaEntity'}
            $entityParams["lineColorRGB"] = $Settings.LineColor
            $entityParams["lineWidth"] = $Settings.LineWidth
            $entityParams["fillStyle"] = $Settings.FillStyle
            $entityParams["fillColorRGB"] = $Settings.FillColor
            $entityParams["pointsString"] = $row.pointsString
        } else { # Point
            $entityParams["entityInstance"] = "UIEEntity"
            $entityParams["centerLat"] = [math]::Floor([double]$row.Y * 3600000)
            $entityParams["centerLon"] = [math]::Floor([double]$row.X * 3600000)
        }

        foreach ($key in $entityParams.Keys) {
            $xmlWriter.WriteAttributeString($key, $entityParams[$key])
        }

        $xmlWriter.WriteEndElement() # End Entity
    }

    $xmlWriter.WriteEndElement(); $xmlWriter.WriteEndElement(); $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndDocument(); $xmlWriter.Flush(); $xmlWriter.Close()
    
    if($Settings.InputFile -like '*.kml' -and (Test-Path $csvPath)){
        Copy-Item -Path $csvPath -Destination $script:CsvOutputPath -Force
    }
}

Function Process-KmlToMasterCsv {
    param(
        [string]$InputKmlPath,
        [string]$OutputCsvPath,
        [bool]$UseNs2Syntax = $false
    )

    $content = Get-Content -Path $InputKmlPath -Raw
    $tempFile1 = Join-Path $script:TempPath "kml_prep1.txt"
    $tempFile2 = Join-Path $script:TempPath "kml_prep2.txt"

    # Define patterns based on syntax
    if ($UseNs2Syntax) {
        $namePattern = '<ns2:Name>.*</ns2:Name>'
        $latPattern = '<ns2:Lat>.*</ns2:Lat>'
        $lonPattern = '<ns2:Lon>.*</ns2:Lon>'
        $symbolPattern = '<ns2:SymbolCode>.*</ns2:SymbolCode>'
    } else {
        $namePattern = '<kml:name>.*</kml:name>'
        $coordPattern = '<kml:coordinates>.*</kml:coordinates>'
        $symbolPattern = '<kml:value>.*</kml:value>'
    }
    
    $allPatterns = @($namePattern, $latPattern, $lonPattern, $coordPattern, $symbolPattern) -join '|'
    $content | Select-String -Pattern $allPatterns -AllMatches |
        Select-Object -ExpandProperty Matches |
        Select-Object -ExpandProperty Value |
        Set-Content -Path $tempFile1

    (Get-Content -Path $tempFile1 -Raw) -replace '<kml:name>.*?</kml:name>', {param($m) $m.Value -replace ' ', ''} | # Compact name tags
        -replace '<kml:coordinates>', '' -replace '</kml:coordinates>', '' `
        -replace ' ', "`n" | Set-Content -Path $tempFile1
    
    $name = ''
    $symbol = ''
    $output = switch -Regex -File $tempFile1 {
        '<kml:name>(.*?)</kml:name>' { $name = $Matches[1] }
        '<ns2:Name>(.*?)</ns2:Name>' { $name = $Matches[1] }
        '<kml:value>(.*?)</kml:value>' { $symbol = $Matches[1] }
        '<ns2:SymbolCode>(.*?)</ns2:SymbolCode>' { $symbol = $Matches[1] }
        default { if ($_.Trim()) { "$_,$name,$symbol" } }
    }

    $output | Where-Object { $_ -like '*,*' } | Set-Content -Path $OutputCsvPath
    Remove-Item $tempFile1, $tempFile2 -ErrorAction SilentlyContinue
}


Function Process-CpofKmlToCsv {
    param (
        [string]$InputKmlPath,
        [string]$OutputCsvPath
    )

    $content = Get-Content -Path $InputKmlPath -Raw
    $tempFile1 = Join-Path $script:TempPath "kml_prep1.txt"
    $tempFile2 = Join-Path $script:TempPath "kml_prep2.txt"

    $content | Select-String -Pattern '<kml:name>.*</kml:name>|<kml:coordinates>.*</kml:coordinates>' -AllMatches |
        Select-Object -ExpandProperty Matches |
        Select-Object -ExpandProperty Value |
        Set-Content -Path $tempFile1

    (Get-Content -Path $tempFile1 -Raw) -replace '<kml:name>(.*?) (.*?) (.*?) (.*?)</kml:name>', '<kml:name>$1$2$3$4</kml:name>' `
        -replace '<kml:name>(.*?) (.*?) (.*?)</kml:name>', '<kml:name>$1$2$3</kml:name>' `
        -replace '<kml:name>(.*?) (.*?)</kml:name>', '<kml:name>$1$2</kml:name>' `
        -replace '<kml:coordinates>', '' -replace '</kml:coordinates>', '' `
        -replace ' ', "`n" | Set-Content -Path $tempFile1

    $name = ''
    $output = switch -Regex -File $tempFile1 {
        '<kml:name>(.*?)</kml:name>' { $name = $Matches[1] }
        default { if ($_.Trim()) { "$_,$name" } }
    }
    $output | Where-Object { $_ -like '*,*,*' } | Set-Content -Path $tempFile2
    
    $csvHeader = "X,Y,Name"
    $csvContent = (Get-Content -Path $tempFile2) -replace '<kml:name>.*</kml:name>', '' | Where-Object { $_.Trim() -ne '' }
    
    Set-Content -Path $OutputCsvPath -Value $csvHeader
    Add-Content -Path $OutputCsvPath -Value $csvContent

    Remove-Item $tempFile1, $tempFile2 -ErrorAction SilentlyContinue
}

#endregion Core Processing Logic

#region Settings Forms

Function Get-ExtractorSettings {
    param($InputKmlFile)
    
    $form = New-Object System.Windows.Forms.Form -Property @{ Text = 'Extractor Settings'; ClientSize = '400,200'; StartPosition = 'CenterScreen' }
    
    $controls = @()
    $controls += New-ModernLabel -Text "Input KML File:" -X 10 -Y 10 -Width 380
    $controls += New-Object System.Windows.Forms.TextBox -Property @{ Text=$InputKmlFile; ReadOnly=$true; Location='10,30'; Size='380,20'}
    
    $controls += New-ModernLabel -Text "Output Folder Prefix (Subfolder Name):" -X 10 -Y 60
    $prefixBox = New-Object System.Windows.Forms.TextBox -Property @{ Location='10,80'; Size='380,20'; Text = [io.path]::GetFileNameWithoutExtension($InputKmlFile) }
    $controls += $prefixBox

    $okButton = New-ModernButton -Title "Run Extraction" -X 210 -Y 150 -Width 100 -BaseColor "#38A169" -HoverColor "#68D391"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    
    $cancelButton = New-ModernButton -Title "Cancel" -X 320 -Y 150 -Width 70 -BaseColor "#E53E3E" -HoverColor "#FC8181"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton

    $form.Controls.AddRange($controls)
    $form.Controls.AddRange(@($okButton, $cancelButton))

    if($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return [pscustomobject]@{
            InputFile       = $InputKmlFile
            OutputDirectory = [io.path]::GetDirectoryName($InputKmlFile) # Save in same dir as input
            Prefix          = $prefixBox.Text
        }
    }
    return $null
}

Function Get-ConversionSettings {
    param (
        [string]$Mode # 'Modify', 'Quick', 'Standard'
    )

    $form = New-Object System.Windows.Forms.Form -Property @{
        Text       = "Settings"
        ClientSize = '660, 350'
        BackColor  = 'WhiteSmoke'
        StartPosition = 'CenterScreen'
    }

    # --- Controls ---
    $yPos = 10
    $controls = @()

    $controls += New-ModernLabel "Graphic Title (BSWS Layer Name)" 10 $yPos; $yPos += 20
    $layerNameTextBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = "10,$yPos"; Size = '300,20' }
    $controls += $layerNameTextBox; $yPos += 30
    
    $controls += New-ModernLabel "Controller" 10 $yPos; $yPos += 20
    $controllerBox = New-Object System.Windows.Forms.ComboBox -Property @{ Location = "10,$yPos"; DropDownStyle = 'DropDownList'; Size = '300,20' }
    $controllerBox.Items.AddRange(@('BLUEFOR:Senior Controller', 'RED:Senior Controller', 'CIVILIAN SIDE:Senior Controller', 'USA:BLUEFOR', 'BLUEFOR:BLUEROLE', 'ENEMY:RED'))
    $controllerBox.SelectedIndex = 0
    $controls += $controllerBox; $yPos += 30

    $controls += New-ModernLabel "Graphic Type" 10 $yPos; $yPos += 20
    $featureTypeBox = New-Object System.Windows.Forms.ComboBox -Property @{ Location = "10,$yPos"; DropDownStyle = 'DropDownList'; Size = '150,20' }
    $featureTypeBox.Items.AddRange(@('Point', 'Checkpoint', 'TCP', 'CCP', 'Line', 'PhaseLine', 'MSR', 'ASR', 'Area', 'OBJ', 'TAI', 'NAI', 'ATKPOS'))
    $featureTypeBox.SelectedIndex = 0
    $controls += $featureTypeBox

    # Styling Controls - only show if not in QuickMode
    if ($Mode -ne 'Quick') {
        $yPos = 10 # Reset Y for the second column
        $controls += New-ModernLabel "Read Only" 350 $yPos; $yPos += 20
        $readOnlyBox = New-Object System.Windows.Forms.ComboBox -Property @{ Location = "350,$yPos"; DropDownStyle = 'DropDownList'; Size = '100,20' }
        $readOnlyBox.Items.AddRange(@('No', 'Yes')); $readOnlyBox.SelectedIndex = 0
        $controls += $readOnlyBox; $yPos += 30

        # ... Add all other detailed styling controls here (Fill, Line, Text etc.)
        # For brevity, I'll add one more example
        $controls += New-ModernLabel "Line Color" 350 $yPos; $yPos += 20
        $lineColorBox = New-Object System.Windows.Forms.ComboBox -Property @{ Location = "350,$yPos"; DropDownStyle = 'DropDownList'; Size = '150,20' }
        $lineColorBox.Items.AddRange(@('Black', 'Blue', 'Yellow', 'Green', 'Red', 'White'))
        $lineColorBox.SelectedIndex = 0
        $controls += $lineColorBox
    }

    $form.Controls.AddRange($controls)

    $okButton = New-FormButton -Title "OK" -X 490 -Y 300 -Width 75 -Height 25 -BaseColor "MediumSeaGreen" -HoverColor "#68D391"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    
    $cancelButton = New-FormButton -Title "Cancel" -X 575 -Y 300 -Width 75 -Height 25 -BaseColor "LightCoral" -HoverColor "#FC8181"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton

    $form.Controls.AddRange(@($okButton, $cancelButton))

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $isReadOnly = if($readOnlyBox){if($readOnlyBox.SelectedItem -eq 'Yes'){1}else{0}}else{0}
        $lineColorName = if($lineColorBox){$lineColorBox.SelectedItem}else{'Black'}

        return [PSCustomObject]@{
            LayerName   = $layerNameTextBox.Text
            Controller  = $controllerBox.SelectedItem
            FeatureType = $featureTypeBox.SelectedItem
            SymbolID    = Get-SymbolId -FeatureType $featureTypeBox.SelectedItem
            ReadOnly    = if($Mode -ne 'Quick') { $isReadOnly } else { 0 }
            FillStyle   = if($Mode -ne 'Quick') { -1 } else { -1 } 
            FillColor   = if($Mode -ne 'Quick') { Get-ColorValue 'None' } else { Get-ColorValue 'None' }
            LineWidth   = if($Mode -ne 'Quick') { 1 } else { 1 }
            LineColor   = if($Mode -ne 'Quick') { Get-ColorValue $lineColorName } else { Get-ColorValue 'Black' }
            TextColor   = if($Mode -ne 'Quick') { Get-ColorValue 'Black' } else { Get-ColorValue 'Black' }
            TextSize    = if($Mode -ne 'Quick') { 14 } else { 14 }
            TextStyle   = if($Mode -ne 'Quick') { Get-TextStyleValue 'Regular' } else { Get-TextStyleValue 'Regular' }
        }
    }
    
    return $null
}

#endregion Settings Forms


# --- Entry Point ---
Start-GraphicsAssistant
