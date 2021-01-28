# Updating Time Series Insights model for Industrial IoT

See `TSIModel-iiot-sample.xlsx` as an example updated model file for instances created from [Azure Industrial IoT Platform](https://github.com/Azure/Industrial-IoT). Updated/added cells in each worksheet is highlighted with yellow background color. An extra worksheet (Mappings) is added to help mapping id codes in original data to descriptive names.

In an industrial IoT scenario where data is originated from OPC UA servers, you may end up with coded values for data source names and tags. [Azure Industrial IoT Platform](https://github.com/Azure/Industrial-IoT) uses 3 fields as Time Series ID: **PublisherId, DataSetWriterId, NodeId**. From those 3 attributes, our aim is to produce consistent Equipment/DataPoint mappings we can use in naming instances and in hierarchies.

* PublisherId is related to the data source (OPC UA Server) and is assumed to map to equipment which generates data.
* DataSetWriterId is related to specifics of data acquisition (subscriptions, publish frequency) and is not used in mapping.
* NodeId is related to actual data point (tag) received from data source.

> Note depending on your configuration, for example, if you connect to multiple OPC Servers/Equipment from a single OPC Publisher module, you can use both PublisherId and DataSetWriterId to map to equipment.

In **Mappings** worksheet you have a table to map PublisherId to Equipment

| PublisherId | Equipment Name |
| ---- | ---- |
|uat124001d2387ce9b3df63d48ec5d5f378c9ab9ec1|USB1GENR01|
|uat29b2f5109d99558d6d13a264c55520eae4f79752|USB1COMP01|
|uat2a9e12b05d99079b713b8db0d3fde8c4c264cbab|USB2GENR01|
|uat312c62d4e0eeeae0689cb09146b2412869a3e850|USB2COMP01|
|uat4c8d6c640f3e9e14f435eb33f5ed69b88bf77725|UKB1GENR01|
|uatcf5ba2465af8de967342c7bd33144ae745b64543|UKB1COMP01|

------

and another table to map NodeId to Datapoint

|NodeId|Datapoint Name|
| ---- | ---- |
|http://microsoft.com/Opc/OpcPlc/#s=AlternatingBoolean|VoltageAlert|
|http://microsoft.com/Opc/OpcPlc/#s=BadFastUInt1|RotationSpeed|
|http://microsoft.com/Opc/OpcPlc/#s=BadSlowUInt1|Vibrations|
|http://microsoft.com/Opc/OpcPlc/#s=DipData|CutPower|
|http://microsoft.com/Opc/OpcPlc/#s=FastUInt1|Voltage|
|http://microsoft.com/Opc/OpcPlc/#s=NegativeTrendData|OilTemperature|
|http://microsoft.com/Opc/OpcPlc/#s=PositiveTrendData|Suction Pressure|
|http://microsoft.com/Opc/OpcPlc/#s=RandomSignedInt32|Power|
|http://microsoft.com/Opc/OpcPlc/#s=RandomUnsignedInt32|OilLevel|
|http://microsoft.com/Opc/OpcPlc/#s=SlowUInt1|DischargePressure|
|http://microsoft.com/Opc/OpcPlc/#s=SpikeData|OilPressure|
|http://microsoft.com/Opc/OpcPlc/#s=StepUp|AirCooledStep|

> Note that Equipment names conform to CCBBTTTT00 format where
> * CC: Site Code, 2 chars (e.g US, UK)
> * BB: Building Code, 2 chars (e.g. B1, B2)
> * TTTT: Equipment type code, 4 chars (e.g. GENR: Generator, COMP: Compressor)
> * 00: Equipment sequence code, 2 digits (e.g. 01, 02, 03)

* **name** field in all **Instances** and **Instances  (\<Hierarchy name\>)** worksheets use these two tables to lookup and join value from both tables to generate Instance name (e.g. USB1GENR01:OilPressure) which uniquely identifies the datapoint.

* Other mapping tables for Functional Location Hierarchy and Asset Type mapping define equipment attributes, which are then used to populate Hierarchy instance fields in **Instances  (\<Hierarchy name\>)** worksheets.

* In **Instances (By Functional Locati** worksheet, all instances already existed in rows (because all instances were already belonged to this hierarcy when we exported data). Therefore, we only updated **name** field and instance fields related to **By Functional Location** hierarchy (which are yellow background columns at the far right) using lookup from mappings.

* In **Instances (By Asset Type)** worksheet, only 14 instances were assigned to **By Asset Type** hierarchy when we exported instances from Time Series Insights. In order to add remaining instances to hierarchy we added their IDs into 3 columns (timeSeriesId1,timeSeriesId2,timeSeriesId3), you can see them starting from row 16. You dont have to fill all columns for these fields, however you need to add **hierarchyId** for each row. Finally we went ahead and updated **name** field and instance fields related to **By Asset Type** hierarchy (which are yellow background columns at the far right) using lookup from mappings.

* **Instances** worksheet contains a list of all instances along with any instance fields which are **NOT** related to any hierarchy. Instance fields for each hierarchy can be found in related **Instances  (\<Hierarchy name\>)** worksheets. Here, we update only **name** and **MeasureUnit** fields.

