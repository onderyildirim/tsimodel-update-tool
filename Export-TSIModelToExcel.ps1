[CmdletBinding()]
Param(
    [Parameter()]
    [Alias("h")]
    [switch]$Help,

    [Parameter(mandatory=$false)]
    [string]$ModelFile=".\TSIModel.xlsx",

    [Parameter(mandatory=$false)]
    [string]$InstancesFile=".\instances.json",

    [Parameter(mandatory=$false)]
    [string]$HierarchiesFile=".\hierarchies.json",

    [Parameter(mandatory=$false)]
    [string]$TypesFile=".\types.json"
    )

function PutGridlines([object]$ws)
{
    $xlEdgeLeft=7
    $xlEdgeTop=8
    $xlEdgeBottom=9
    $xlEdgeRight=10
    $xlInsideVertical=11
    $xlInsideHorizontal=12
    $xlContinuous=1
    $ws.UsedRange.Borders($xlEdgeTop).LineStyle = $xlContinuous
    $ws.UsedRange.Borders($xlEdgeLeft).LineStyle = $xlContinuous
    $ws.UsedRange.Borders($xlEdgeBottom).LineStyle = $xlContinuous
    $ws.UsedRange.Borders($xlEdgeRight).LineStyle = $xlContinuous
    $ws.UsedRange.Borders($xlInsideVertical).LineStyle = $xlContinuous
    $ws.UsedRange.Borders($xlInsideHorizontal).LineStyle = $xlContinuous
}

function MarkExcelColumnUpdateable([object]$ws,[int]$wscol)
{
    $wsaddress  = $ws.Cells(1,$wscol).AddressLocal()
    $wsaddress += ":" 
    $wsaddress += $ws.Cells($ws.UsedRange.Rows.Count,$wscol).AddressLocal()

    MarkExcelRangeUpdateable $ws.Range($wsaddress)
}

function MarkExcelRangeUpdateable([object]$range)
{
    $xlSolid=1
    $xlAutomatic=-4105
    $xlThemeColorAccent4=8
    $range.Interior.Color = 16773836
    $range.Interior.Pattern = $xlSolid
    $range.Interior.PatternColorIndex = $xlAutomatic
    $range.Interior.ThemeColor = $xlThemeColorAccent4
    $range.Interior.TintAndShade = 0.799981688894314
    $range.Interior.PatternTintAndShade = 0
}


if($Help -eq $true)
{
    Write-Output "Usage: $($MyInvocation.MyCommand.Name) [OPTIONS]"
    Write-Output "OPTIONS:"
    Write-Output "   -h, Help        : Display this screen."
    Write-Output "   -InstancesFile  : Instances file exported from TSI. Default is 'instances.json'."
    Write-Output "   -HierarchiesFile: Hierarchies file exported from TSI. Default is 'hierarchies.json'."
    Write-Output "   -TypesFile      : Types file exported from TSI. Default is 'types.json'."
    Write-Output "   -ModelFile      : Path to the output Excel file, to be modified and fed into 'Import-TSIModelFromExcel'. Default is 'TSIModel.xlsx'."
    Write-Output ""
    Write-Output "$($MyInvocation.MyCommand.Name) -ModelFile `".\TSIModel.xlsx`" -InstancesFile `".\instances.json`" -HierarchiesFile `".\hierarchies.json`" -TypesFile `".\types.json`""
    Exit 0
}


if (-not (Test-Path -Path $InstancesFile))
{
    Write-Output "Instances file '$InstancesFile' not found."
    Exit 1
}

if (-not (Test-Path -Path $HierarchiesFile))
{
    Write-Output "Hierarchies file '$HierarchiesFile' not found."
    Exit 1
}

if (-not (Test-Path -Path $TypesFile))
{
    Write-Output "Types file '$TypesFile' not found."
    Exit 1
}


if (Test-Path -Path $ModelFile)
{
    If((Read-Host -Prompt "'$ModelFile' exists and will be overwritten do you want to continue? (y/n)").ToLowerInvariant() -ne 'y')
        {exit}
}

$jsonInstances   = (Get-Content -Raw -Path $InstancesFile   | ConvertFrom-Json).put
$jsonHierarchies = (Get-Content -Raw -Path $HierarchiesFile | ConvertFrom-Json).put
$jsonTypes       = (Get-Content -Raw -Path $TypesFile       | ConvertFrom-Json).put

Write-Output "Opening Excel..."
$XL = New-Object -comobject Excel.Application
$XL.Visible = $False
$XL.DisplayAlerts = $false
$wb = $XL.Workbooks.Add()

Write-Output "Exporting types..."
$typesWS=$wb.Worksheets.Item(1)
$typesWS.Name = "Types"
$types = $jsonTypes | sort -Property id | ForEach-Object {[pscustomobject]@{ id=$_.id; name=$_.name; desc=$_.description; vars=($_.variables.PSObject.Properties.Name -join ',')}}
$typesWS.cells.item(1,1) = "id"
$typesWS.cells.item(1,2) = "name"
$typesWS.cells.item(1,3) = "description"
$typesWS.cells.item(1,4) = "variables"

$line=2
foreach($type in $types)
{
    $typesWS.cells.item($line,1) = $type.id
    $typesWS.cells.item($line,2) = $type.name
    $typesWS.cells.item($line,3) = $type.desc
    $typesWS.cells.item($line,4) = $type.vars
    $line=$line+1
}
$typesWS.Columns.AutoFit() | Out-Null
PutGridlines $typesWS

Write-Output "Exporting hierarchies..."
$hierarchiesWS=$wb.Worksheets.Add()
$hierarchiesWS.Name = "Hierarchies"
$hierarchies = $jsonHierarchies| sort -Property id | ForEach-Object {[pscustomobject]@{ id=$_.id; name=$_.name; fields=($_.source.instanceFieldNames -join ','); instanceFieldNames=$_.source.instanceFieldNames}}
$hierarchiesWS.cells.item(1,1) = "id"
$hierarchiesWS.cells.item(1,2) = "name"
$hierarchiesWS.cells.item(1,3) = "instanceFields"

$line=2
foreach($h in $hierarchies)
{
    $hierarchiesWS.cells.item($line,1) = $h.id
    $hierarchiesWS.cells.item($line,2) = $h.name
    $hierarchiesWS.cells.item($line,3) = $h.fields
    $line=$line+1
}

$hierarchiesWS.Columns.AutoFit() | Out-Null
PutGridlines $hierarchiesWS

Write-Output "Exporting instances..."
$instancesWS=$wb.Worksheets.Add()
$instancesWS.Name = "Instances"
$instances = $jsonInstances | ForEach-Object {[pscustomobject]@{ typeid=$_.typeId; name=$_.name; id=($_.timeSeriesId -join ','); timeSeriesId=$_.timeSeriesId; instanceFields=$_.instanceFields; hierarchyIds=$_.hierarchyIds}}
$colNum=1
$line=1
$instancesWS.cells.item($line,$colNum++) = "typeId"
$instancesWS.cells.item($line,$colNum++) = "typeName"

if($jsonInstances[0].timeSeriesId.Length -eq 1)
{
    $instancesWS.cells.item($line,$colNum++) = "timeSeriesId"
}
else
{
    foreach($t in $jsonInstances[0].timeSeriesId)
    {
        $instancesWS.cells.item($line,$colNum++) = "timeSeriesId"+ ($colNum-3).ToString()
    }
}

$instancesWS.cells.item($line,$colNum++)= "name"

$lastColIndex=$colNum
$nonHierarchyInstanceFields=@{}

$line++
foreach($in in $instances)
{
    $colNum=1
    $instancesWS.cells.item($line,$colNum++) = $in.typeId
    $instancesWS.cells.item($line,$colNum++) = ($jsonTypes|where id -eq $in.typeId).name
    foreach($t in $in.timeSeriesId)
    {
        $instancesWS.cells.item($line,$colNum++) = $t
    }
    $instancesWS.cells.item($line,$colNum++) = $in.name

    $fieldList=[ordered]@{}
    $in.instanceFields.PSObject.Properties.ForEach({ $fieldList[$_.Name] = $_.Value })

    foreach($ih in $in.hierarchyIds)
    {
        $hfields = $jsonHierarchies| where id -eq $ih | select -ExpandProperty source | select -ExpandProperty instanceFieldNames
        foreach($hf in $hfields)
        {
            $fieldList.Remove($hf)
        }
    }

    foreach($f in $fieldList.Keys)
    {
        
        if($nonHierarchyInstanceFields.ContainsKey($f))
        {
            $ifColNum=$nonHierarchyInstanceFields[$f]
        }
        else
        {
            $ifColNum=$lastColIndex++
            $nonHierarchyInstanceFields[$f]=$ifColNum
            $instancesWS.cells.item(1,$ifColNum)=$f
        }

        $instancesWS.cells.item($line,$ifColNum)=$fieldList.$f
        MarkExcelRangeUpdateable $instancesWS.cells($line,$ifColNum)

    }


    $line=$line+1
    $pct =[int] ((($line-2)/$instances.Count)*100)
    Write-Progress -Activity "Exporting instances..." -Status "$pct% ($($line-2)/$($instances.Count)) Complete:" -PercentComplete $pct
}


$instancesWS.Columns.AutoFit() | Out-Null
PutGridlines $instancesWS
MarkExcelColumnUpdateable $instancesWS 1
MarkExcelColumnUpdateable $instancesWS $($lastColIndex-1)

foreach($h in $hierarchies)
{
    Write-Output "Exporting instances for hierarchy '$($h.name)'..."
    $instancesWS=$wb.Worksheets.Add()
    $wsName="Instances ($($h.name))"
    if ($wsName.Length -gt 31) {$wsName=$wsName.Substring(0, 31)}
    $instancesWS.Name = $wsName
    $instances = $jsonInstances | ForEach-Object {[pscustomobject]@{ typeid=$_.typeId; name=$_.name; id=($_.timeSeriesId -join ','); timeSeriesId=$_.timeSeriesId; instanceFields=$_.instanceFields; hierarchyIds=$_.hierarchyIds}} | where hierarchyIds -CContains $h.id
    $colNum=1
    $line=1
    $instancesWS.cells.item($line,$colNum++) = "typeId"
    $instancesWS.cells.item($line,$colNum++) = "typeName"

    if($jsonInstances[0].timeSeriesId.Length -eq 1)
    {
        $instancesWS.cells.item($line,$colNum++) = "timeSeriesId"
    }
    else
    {
        foreach($t in $jsonInstances[0].timeSeriesId)
        {
            $instancesWS.cells.item($line,$colNum++) = "timeSeriesId"+ ($colNum-3).ToString()
        }
    }

    $instancesWS.cells.item($line,$colNum++)= "name"
    $instancesWS.cells.item($line,$colNum++)= "hierarchyId"
    $instancesWS.cells.item($line,$colNum++)= "hierarchyName"
    $startOfInstanceFields=$colNum
    foreach($ifn in $h.instanceFieldNames)
    {
        $instancesWS.cells.item($line,$colNum++) = $ifn
    }

    $line++
    foreach($in in $instances)
    {
        $colNum=1
        $instancesWS.cells.item($line,$colNum++) = $in.typeId
        $instancesWS.cells.item($line,$colNum++) = ($jsonTypes|where id -eq $in.typeId).name
        foreach($t in $in.timeSeriesId)
        {
            $instancesWS.cells.item($line,$colNum++) = $t
        }
        $instancesWS.cells.item($line,$colNum++) = $in.name
        $instancesWS.cells.item($line,$colNum++) = $h.id
        $instancesWS.cells.item($line,$colNum++) = $h.name

        foreach($t in $h.instanceFieldNames)
        {
            if (($in.instanceFields) -and ($in.instanceFields.PSObject.Properties))
            {
                $instancesWS.cells.item($line,$colNum++) = $in.instanceFields.PSObject.Properties[$t].Value
            }
        }

        $line=$line+1
        $pct =[int] ((($line-2)/$instances.Count)*100)
        Write-Progress -Activity "Exporting instances for hierarchy '$($h.name)'..." -Status "$pct% ($($line-2)/$($instances.Count)) Complete:" -PercentComplete $pct

    }
    $instancesWS.Columns.AutoFit() | Out-Null
    PutGridlines $instancesWS
    MarkExcelColumnUpdateable $instancesWS 1
    $i=$startOfInstanceFields
    while($i -le $instancesWS.UsedRange.Columns.Count)
    {
        MarkExcelColumnUpdateable $instancesWS $i
        $i=$i+1
    }

}

Write-Output "Writing to file: $ModelFile..."
$xlWorkbookDefault=51
$wb.SaveAs($ModelFile,$xlWorkbookDefault)

Write-Output "Cleanup..."
$wb.Close($false)
$XL.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($XL) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Remove-Variable XL -ErrorAction SilentlyContinue

Write-Output "Done."


