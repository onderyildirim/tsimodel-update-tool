﻿[CmdletBinding()]
Param(
    [Parameter()]
    [Alias("h")]
    [switch]$Help,

    [Parameter(mandatory=$false)]
    [string]$ModelFile=".\TSIModel.xlsx",

    [Parameter(mandatory=$false)]
    [string]$InstancesFile=".\instances_out.json"
    )

function FormatDuration()
{
Param ([TimeSpan]$TimeSpan)
    $TimeSpanStr = ""
    if($TimeSpan.Days>0) {$TimeSpanStr = [String]::Format("{0} days ",$TimeSpan.Days)}
    $TimeSpanStr = [String]::Format("{0:hh}:{0:mm}:{0:ss}",$TimeSpan)
    return $TimeSpanStr
}

if($Help -eq $true)
{
    Write-Output "Usage: $($MyInvocation.MyCommand.Name) [OPTIONS]"
    Write-Output "OPTIONS:"
    Write-Output "   -h, Help        : Display this screen."
    Write-Output "   -InstancesFile  : Instances file created to be imported back into TSI. Default is 'instances_out.json'."
    Write-Output "   -ModelFile      : Path to the input Excel file, created by 'Export-TSIModelToExcel'. Default is 'TSIModel.xlsx'."
    Write-Output ""
    Write-Output "$($MyInvocation.MyCommand.Name) -InstancesFile `".\instances_out.json`" -ModelFile `".\TSIModel.xlsx`""
    Exit 0
}

$path=Split-Path -Path $ModelFile
if(([string]::IsNullOrEmpty($path)) -or ($path -eq ".")){$ModelFile=$PSScriptRoot+"\"+(Split-Path -Path $ModelFile -Leaf)}

if (-not (Test-Path -Path $ModelFile))
{
    Write-Output "Model file '$ModelFile' not found."
    Exit 1
}


if (Test-Path -Path $InstancesFile)
{
    If((Read-Host -Prompt "'$InstancesFile' exists and will be overwritten do you want to continue? (y/n)").ToLowerInvariant() -eq 'y')
        {Remove-Item $InstancesFile}
    else
        {exit}
}


Write-Output "Opening Excel..."
$XL = New-Object -comobject Excel.Application
$XL.Visible = $False
$wb = $XL.Workbooks.Open($ModelFile, $False, $True)
$instancesJson = [System.Collections.ArrayList][ordered]@{}


foreach($ws in $wb.Worksheets)
{
    $timeStarted= Get-Date
    if ($ws.name -like "Instances*")
    {
        Write-Output "Processing Sheet: $($ws.name)..."

        $line=2
        $tsiIdStartColumn=3
        $wsRowCount = $ws.UsedRange.Rows.Count


        While($ws.cells.item($line,$tsiIdStartColumn).Value())
        {
            $colNum=$tsiIdStartColumn

            $timeSeriesId=[System.Collections.ArrayList]@()
            $tsidNumCols=0
            while($colNum -lt 6)
            {
                $h = $ws.cells.item(1,$colNum).Value()
                if($h  -like "timeSeriesId*")
                {
                    [void]$timeSeriesId.Add($ws.cells.item($line,$colNum++).Value())
                    $tsidNumCols++
                }
                else
                {
                    break
                }
            }
            $currentNode = $instancesJson | where {(Compare-Object $_.timeSeriesId $timeSeriesId -ExcludeDifferent -IncludeEqual).InputObject.Count -eq $tsidNumCols}
            if (-not $currentNode)
            {
                $currentNode=[ordered]@{'typeId'=$ws.cells.item($line,1).Value(); 'timeSeriesId'=$timeSeriesId; }
                [void]$instancesJson.Add($currentNode)
            }
            else
            {
                if ($ws.cells.item($line,1).Value())
                {
                    $currentNode.typeId=$ws.cells.item($line,1).Value()
                }
            }

            if($ws.cells.item($line,$colNum).Value()){$currentNode.name=$ws.cells.item($line,$colNum).Value()}

            if ($ws.name -eq "Instances")
            {
                $colNum=$colNum+1
                if($ws.cells.item($line,$colNum).Value()){$currentNode.description=$ws.cells.item($line,$colNum).Value()}
            }

            if($ws.cells.item(1,$colNum).Value())
            {
                $colNum=$colNum+1
                if($ws.cells.item(1,$colNum).Value() -eq "hierarchyId")
                {
                    if($ws.cells.item(1,$colNum).Value())
                    {
                        if(-not $currentNode.hierarchyIds)
                        {
                            $currentNode.hierarchyIds=[System.Collections.ArrayList]@()
                        }
                        $hnode=$currentNode.hierarchyIds
                        [void]$hnode.Add($ws.cells.item($line,$colNum).Value())
                        $colNum=$colNum+1
                        $colNum=$colNum+1

                        if(-not $currentNode.instanceFields)
                        {
                            $currentNode.instanceFields=[ordered]@{}
                        }
                        $inode=$currentNode.instanceFields
                        while($ws.cells.item(1,$colNum).Value())
                        {
                            if ($ws.cells.item($line,$colNum).Value())
                            {
                                if(-not $inode.Contains($ws.cells.item(1,$colNum).Value()))
                                {
                                    [void]$inode.Add($ws.cells.item(1,$colNum).Value(),$ws.cells.item($line,$colNum).Value())
                                }
                                else
                                {
                                    $inode[$ws.cells.item(1,$colNum).Value()]=$ws.cells.item($line,$colNum).Value()
                                }
                            }
                            $colNum=$colNum+1
                        }
                    }
                    else
                    {
                        Write-Host "[warn ] Worksheet: $($ws.name), Line: $line : hierarchyId is empty. Skipped record." -ForegroundColor Yellow
                    }
                }
                else
                {
                    $i=$colNum
                    if((-not $currentNode.instanceFields) -and ($ws.cells.item(1,$i).Value()))
                    {
                        $currentNode.instanceFields=[ordered]@{}
                    }
                    while($ws.cells.item(1,$i).Value())
                    {
                        $instanceFieldName=$ws.cells.item(1,$i).Value()
                        $instanceFieldValue=$ws.cells.item($line,$i).Value()
                        if($instanceFieldValue)
                        {
                            $currentNode.instanceFields.$instanceFieldName=$instanceFieldValue
                        }
                        $i=$i+1
                    }
                }
            }
            
            $line=$line+1

            $currentItem=$line
            $totalItems=$wsRowCount
            if (($currentItem -le $totalItems) -and ($totalItems -le 1000) -or (($totalItems -gt 1000) -and ($currentItem%100 -eq 0)))
            {
                $completeratio = [double](($currentItem)/[double]$totalItems)
                $pct =[int] ($completeratio*100)
                if($pct -gt 100){$pct=100}
                $timeElapsed = New-TimeSpan -Start $timeStarted -End $(Get-Date)
                $totalTime = New-TimeSpan -Seconds ([Int64]([double]$timeElapsed.TotalSeconds * [double]$totalItems / [double]($currentItem)))
                $timeRemaining=$totalTime - $timeElapsed

                $timeRemainingStr = FormatDuration -TimeSpan $timeRemaining
                $timeElapsedStr = FormatDuration -TimeSpan $timeElapsed
                $totalTimeStr = FormatDuration -TimeSpan $totalTime

                $statusMsg=[String]::Format("{0}% ({1}/{2}) Complete. Time Elapsed:{3} Time Remaining:{4} Total Time:{5}", $pct,$currentItem,$totalItems,$timeElapsedStr, $timeRemainingStr, $totalTimeStr)
                Write-Progress -Activity "Processing Sheet: $($ws.name)..." -Status $statusMsg -PercentComplete $pct
            }

        }

    } 
}

Write-Output "Writing to file: $InstancesFile..."
$instancesJsonTop = [System.Collections.ArrayList][ordered]@{}
$instancesJsonTop=[ordered]@{'put'=$instancesJson}
$instancesJsonTop | ConvertTo-Json -depth 100 | Out-File $InstancesFile 


Write-Output "Cleanup..."
$wb.Close($false)
$XL.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Remove-Variable XL -ErrorAction SilentlyContinue

Write-Output "Done."
