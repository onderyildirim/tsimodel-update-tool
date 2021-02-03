# How to update Time Series Insights model to create an industrial asset framework

Asset Framework is a crucial component of any industrial IoT infrastructure. When you collect your IoT data into TSI, you will need a proper AF model to be able to drive insights from it.

This tool enables users to export and modify TSI model within Excel and deploy results back to TSI to form your own asset framework. 

Two scripts are provided for this purpose:

* `Export-TSIModelToExcel.ps1`  : Exports TSI model data from JSON file(s) into Excel
* `Import-TSIModelFromExcel.ps1`: Imports data from Excel file into JSON file later to be imported into TSI Model.

# Usage

Follow steps below for a complete cycle of updating your model in TSI.

## 1) Requirements

* Excel 2010 or later

* [Powershell v5 or later](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell?view=powershell-7.1)

## 2) Export JSON metadata from Time Series Insights

You first need to export TSI model data from TSI Explorer

* Navigate to TSI Explorer URL
* Click on TSI Model icon
* ![TSI Explorer](./images/tsiexplorer1.png)
* Click on "Instances" tab and "Download instances"
* Save "instances.json" file
* Click on "Hierarchies" tab and "Download hierarchies"
* Save "hierarchies.json" file
* Click on "Types" tab and "Download types"
* Save "types.json" file

## 3) Export model metadata from JSON into Excel format

* Run `Export-TSIModelToExcel.ps1` to export JSON document into Excel (xlsx) format. The script uses Office COM components so you need to have Excel 2010 or later installed in your workstation.

USAGE: `Export-TSIModelToExcel.ps1  [OPTIONS]`

SAMPLE: `Export-TSIModelToExcel.ps1  -InstancesFile instances.json -HierarchiesFile hierarchies.json -TypesFile types.json -ModelFile TSIModel.xlsx`

OPTIONS:<br />
`​          -h, Help        : Display this screen.`<br />
`​          -InstancesFile  : Instances file exported from TSI. Default is 'instances.json'.`<br />
`​          -HierarchiesFile: Hierarchies file exported from TSI. Default is 'hierarchies.json'.`<br />
`​          -TypesFile      : Types file exported from TSI. Default is 'types.json'.`<br />
`​          -ModelFile      : Path to the output Excel file, to be modified and fed into 'Import-TSIModelFromExcel'. Default is 'TSIModel.xlsx'.`<br />


## 4) Update the model in Excel

### 4.1) Background
- Exported model in Excel includes following worksheets
  - `Instances` : Worksheet contains the list of instances, their types and names
  - `Instances (\<Hierarchy name\>)`: Excel workbook contains a separate worksheet for each hierarchy that shows instances instances in that hierarchy
  - `Types`     : List of types. The tool does not update types and this worksheet is only included for reference.
  - `Hirarchies`: List of hierarchies. The tool does not update hierarchies and this worksheet is only included for reference.

- Updatable columns in each worksheet is marked with yellow background.
- **timeSeries** column(s) are the key in each of "Intances*" worksheets.
- **Instances** worksheet has a line for each **timeSeriesId** which needs to be unique.  
- **Instances  (\<Hierarchy name\>)** worksheets exist for each Hierarchy. These worksheets contain typeId/Name, timeSeriesID, hierarchyId/Name and any instance fields for hierarchy.
- Import script processes worksheets in order (left-to-right). If an attribute value (such as typeId or name) is null on any worksheet, existing value is left unchanged. So that when you wan to update those attributes that exist in multiple worksheets you can only update one and leave the rest "null".

### 4.2) Change instance types
Update **typeId** column in any of ** Instances* ** worksheets to change instance type. Tools will preserve last value in processing worksheets left-to-right. 

### 4.3) Add/remove new instances
Add/remove lines into **Instances** worksheet. You can also add lines into any of **Instances  (\<Hierarchy name\>)** worksheets to create an instance and include that in the hierarchy. 
* For an instance to be added, it needs to be in one of **Instances\*** worksheets
* For an instance to be deleted, it needs to be deleted in all **Instances\*** worksheets

### 4.4) Add/remove new instances into hierarchy
Add/remove lines into/from related **Instances  (\<Hierarchy name\>)** worksheets. Note that you have to repeat the same **hierarchyId** in each line added.

### 4.5) Update instances' placement within hierarchy
Modify **instance fields** (rightmost fields within worksheet after **hierarchy name**) within related **Instances  (\<Hierarchy name\>)** worksheet.

### 4.6) Changing Instance Name
**name** property can be set in any **Instances\*** worksheet. Tool will process worksheets from left to right and use the last non null value for **name** property of any instance.

### 4.7) Adding/Changing Instance fields
Update instance fields in any **Instances\*** worksheets. In case there are duplicates, toll will use the last value in left-to-right processing order.

You can add new instance fields that are not related to any hierarchy in **Instances** worksheet.

### 4.8) Null values for instance fields
The tool will not create the instance field in json document if its value is null for any instance.

### 4.9) Duplicate instance fields
If an instance field is used in two hierarchies the it will appear in two worksheets. Since you cannot have two instance fields with the same name in time series instance, when converting into JSON, Tool will process worksheets from left to right and use the last non null value it finds for an instance field.

### 4.10) Update model file to create the asset framework
For an example of updating the model file from your existing asset frameworkin an industrial IoT scenario, see [Updating Time Series Insights model for Industrial IoT](./iiot-mappings.md)

## 5) Import TSI Model data from Excel into JSON
* Run `Export-TSIModelToExcel.ps1` to receive the updated model in Excel file and generate the JSON file to be uploaded to TSI. 

USAGE: `Export-TSIModelToExcel.ps1  [OPTIONS]`

SAMPLE: `Export-TSIModelToExcel.ps1  -InstancesFile instances_out.json -ModelFile TSIModel.xlsx`

OPTIONS:<br />
`​          -h, Help        : Display this screen.`<br />
`​          -InstancesFile  : Instances file created to be imported back into TSI. Default is 'instances_out.json'.`<br />
`​          -ModelFile      : Path to the input Excel file, created by 'Export-TSIModelToExcel'. Default is 'TSIModel.xlsx'.`<br />

## 6) Import TSI Model from JSON into TSI
The final step is to import modified instances json file back into TSI 

* Navigate to TSI Explorer URL
* Click on TSI Model icon
* ![TSI Explorer](./images/tsiexplorer1.png)
* Click on "Instances" tab and "Upload JSON"
* Click "Choose file" and select **instances_out.json** file you created in step 5.
* Click "Upload"
