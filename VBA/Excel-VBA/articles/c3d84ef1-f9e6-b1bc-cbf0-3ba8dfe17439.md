
# PivotCache Object (Excel)

Represents the memory cache for a PivotTable report.


## Remarks

 The **PivotCache** object is a member of the **[PivotCaches](cfd979b9-d52f-f34b-4b66-4fb17efcdc92.md)** collection.


## Example

Use the  **[PivotCache](82602154-783d-3f78-b354-0dabfdc34c98.md)** method to return a **PivotCache** object for a PivotTable report (each report has only one cache). The following example causes the first PivotTable report on the first worksheet to refresh itself whenever its file is opened.


```
Worksheets(1).PivotTables(1).PivotCache.RefreshOnFileOpen = True
```

Use  **[PivotCaches](0a2e7f10-c123-5c98-fb71-56868b9f8bde.md)** ( _index_ ), where _index_ is the PivotTable cache number, to return a single **PivotCache** object from the **PivotCaches** collection for a workbook. The following example refreshes cache one.




```
ActiveWorkbook.PivotCaches(1).Refresh
```


## Methods



|**Name**|
|:-----|
|[CreatePivotChart](5aeb9a16-2cf8-3525-12b0-0b6e3d3ddf1a.md)|
|[CreatePivotTable](dca20930-5d58-8db7-bd81-3c90b7588011.md)|
|[MakeConnection](d0b29374-4d5a-7d9e-630a-500b505da1bd.md)|
|[Refresh](2833d199-342c-9e2e-d1f8-88c33a74bac6.md)|
|[ResetTimer](846a6d82-a86f-ea3c-f0b7-0481bda02470.md)|
|[SaveAsODC](d7b553a5-70b1-41e7-9e35-088c23357570.md)|

## Properties



|**Name**|
|:-----|
|[ADOConnection](410a3eee-0dda-4be1-45c4-809893de624e.md)|
|[Application](da312f38-5253-05b2-f7a4-e1779a8bd90e.md)|
|[BackgroundQuery](91909d27-68ca-a870-5cd9-72019c65f060.md)|
|[CommandText](07921bda-74fe-2a41-15f7-16068ce49a31.md)|
|[CommandType](bbe0ba26-efb9-428d-de2c-576116d92747.md)|
|[Connection](5d4b07f2-dad9-4c90-ec92-094dac95a086.md)|
|[Creator](3393e844-b6e1-f767-d993-53844536782c.md)|
|[EnableRefresh](5919198f-bb4a-eb54-1a28-41033b525fa1.md)|
|[Index](a806f65f-69c5-0691-8a7d-e6a4601116b4.md)|
|[IsConnected](5c238338-c242-019c-1a29-08d2c87bc3be.md)|
|[LocalConnection](3afee878-3c05-6b05-4770-e10e4c6f9375.md)|
|[MaintainConnection](1fba45e7-0059-26d1-1433-631ee08c0dd0.md)|
|[MemoryUsed](f68731ec-053e-79e9-861f-2c225b827e96.md)|
|[MissingItemsLimit](ff15a86c-b57f-ed55-bbfa-74e1c5ce753c.md)|
|[OLAP](d40d3a71-0a27-c4a6-0c3b-47ab7a1a0e06.md)|
|[OptimizeCache](4aedf3bb-e15a-439c-5987-ea16cc233a7c.md)|
|[Parent](b0b2c1c7-56fc-a9ac-418a-d14dc6673d97.md)|
|[QueryType](61346ed2-1ada-a105-1894-b22861047c4f.md)|
|[RecordCount](5fcdcf2d-d52f-6ac1-ef09-8377fc5a1f4d.md)|
|[Recordset](25f2eb4f-d78c-21e2-9d26-c8ebc3404607.md)|
|[RefreshDate](0bbb3e62-584b-7daf-2ad0-643a6e886187.md)|
|[RefreshName](a44a9b7c-3284-a7ca-3cda-99457ce7c1c4.md)|
|[RefreshOnFileOpen](aed513aa-b752-8b6e-0d6d-6fddab46df18.md)|
|[RefreshPeriod](6357769c-e73e-2388-962a-f3bb790c423e.md)|
|[RobustConnect](354d0124-e178-342b-9565-fa74e9dae5d5.md)|
|[SavePassword](6ddc953a-b014-589b-5b67-7497da9df706.md)|
|[SourceConnectionFile](87755bde-3c43-3520-24f7-2c778a225b18.md)|
|[SourceData](5a172543-3a06-9db0-7edc-0cf2aa7af114.md)|
|[SourceDataFile](1b90ee17-45c1-3c96-33e3-ec6c5515d9ee.md)|
|[SourceType](197da621-7407-e95a-2f5b-1cbe0ec403b0.md)|
|[UpgradeOnRefresh](9110a82b-9ac7-3d9e-8386-827cd828aace.md)|
|[UseLocalConnection](ce54adf2-22f3-f4dc-8b97-276d6ca53478.md)|
|[Version](357f61a1-7401-46c1-2a47-4172fb045cd5.md)|
|[WorkbookConnection](cb4de0b8-6706-f1e3-4e2d-42b38b93c601.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)