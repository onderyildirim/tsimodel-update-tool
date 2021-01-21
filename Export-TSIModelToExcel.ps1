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

function MarkExcelColumnForUpdate([object]$ws,[int]$wscol)
{
    $xlSolid=1
    $xlAutomatic=-4105
    $xlThemeColorAccent4=8
    $ws.Cells.Range("A1:A1").Interior.Pattern = $xlSolid
    $wsaddress  = $ws.Cells(1,$wscol).AddressLocal()
    $wsaddress += ":" 
    $wsaddress += $ws.Cells($ws.UsedRange.Rows.Count,$wscol).AddressLocal()
    $ws.Range($wsaddress).Interior.Color = 16773836
    $ws.Range($wsaddress).Interior.Pattern = $xlSolid
    $ws.Range($wsaddress).Interior.PatternColorIndex = $xlAutomatic
    $ws.Range($wsaddress).Interior.ThemeColor = $xlThemeColorAccent4
    $ws.Range($wsaddress).Interior.TintAndShade = 0.799981688894314
    $ws.Range($wsaddress).Interior.PatternTintAndShade = 0
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
    Exit 0
}

$path=Split-Path -Path $ModelFile
if(([string]::IsNullOrEmpty($path)) -or ($path=".")){$ModelFile=$PSScriptRoot+"\"+(Split-Path -Path $ModelFile -Leaf)}

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
$XL.Visible = $True
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
$instances = $jsonInstances | ForEach-Object {[pscustomobject]@{ typeid=$_.typeId; name=$_.name; id=($_.timeSeriesId -join ','); timeSeriesId=$_.timeSeriesId}}
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
    $line=$line+1
}


$instancesWS.Columns.AutoFit() | Out-Null
PutGridlines $instancesWS
MarkExcelColumnForUpdate $instancesWS 1
MarkExcelColumnForUpdate $instancesWS 6

foreach($h in $hierarchies)
{
    Write-Output "Exporting instances for hierarchy '$($h.name)'..."
    $instancesWS=$wb.Worksheets.Add()
    $instancesWS.Name = "Instances ($($h.name))"
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
            $instancesWS.cells.item($line,$colNum++) = $in.instanceFields.PSObject.Properties[$t].Value
        }

        $line=$line+1
    }
    $instancesWS.Columns.AutoFit() | Out-Null
    PutGridlines $instancesWS
    MarkExcelColumnForUpdate $instancesWS 1
    $i=$startOfInstanceFields
    while($i -le $instancesWS.UsedRange.Columns.Count)
    {
        MarkExcelColumnForUpdate $instancesWS $i
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


