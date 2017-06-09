---
title: PivotTable.PivotTableWizard Method (Excel)
keywords: vbaxl10.chm235117
f1_keywords:
- vbaxl10.chm235117
ms.prod: excel
api_name:
- Excel.PivotTable.PivotTableWizard
ms.assetid: 65b988fd-7b0a-ebcb-61ca-8ece8819aebf
ms.date: 06/08/2017
---


# PivotTable.PivotTableWizard Method (Excel)

Creates and returns a  **[PivotTable](pivottable-object-excel.md)** object. This method doesn?t display the PivotTable Wizard. This method isn?t available for OLE DB data sources. Use the **[Add](pivottables-add-method-excel.md)** method to add a PivotTable cache, and then create a PivotTable report based on the cache.


## Syntax

 _expression_ . **PivotTableWizard**( **_SourceType_** , **_SourceData_** , **_TableDestination_** , **_TableName_** , **_RowGrand_** , **_ColumnGrand_** , **_SaveData_** , **_HasAutoFormat_** , **_AutoPage_** , **_Reserved_** , **_BackgroundQuery_** , **_OptimizeCache_** , **_PageFieldOrder_** , **_PageFieldWrapCount_** , **_ReadData_** , **_Connection_** )

 _expression_ A variable that represents a **PivotTable** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SourceType_|Optional| **Variant**|An  **[XlPivotTableSourceType](xlpivottablesourcetype-enumeration-excel.md)** value that represents the source of the report data. If you specify this argument, you must also specify _SourceData_. If  _SourceType_ and _SourceData_ are omitted, Microsoft Excel assumes that the source type is **xlDatabase** , and the source data comes from the named range "Database." If this named range doesn?t exist, Microsoft Excel uses the current region if the current selection is in a range of more than 10 cells that contain data. If this isn?t true, this method will fail.|
| _SourceData_|Optional| **Variant**|The data for the new report. Can be a  **[Range](range-object-excel.md)** object, an array of ranges, or a text constant that represents the name of another report. For an external database, _SourceData_ is an array of strings containing the SQL query string, where each element is up to 255 characters in length. You should use the _Connection_ argument to specify the ODBC connection string. For compatibility with earlier versions of Excel, _SourceData_ can be a two-element array. The first element is the connection string specifying the ODBC source for the data. The second element is the SQL query string used to get the data. If you specify _SourceData_, you must also specify  _SourceType_. If the active cell is inside the  _SourceData_ range, you must specify _TableDestination_ as well.|
| _TableDestination_|Optional| **Variant**|A  **Range** object specifying where the report should be placed on the worksheet. If this argument is omitted, the report is placed at the active cell.|
| _TableName_|Optional| **Variant**|A string that specifies the name of the new report.|
| _RowGrand_|Optional| **Variant**| **True** to show grand totals for rows in the report.|
| _ColumnGrand_|Optional| **Variant**| **True** to show grand totals for columns in the report.|
| _SaveData_|Optional| **Variant**| **True** to save data with the report. **False** to save only the report definition.|
| _HasAutoFormat_|Optional| **Variant**| **True** to have Microsoft Excel automatically format the report when it?s refreshed or when fields are moved.|
| _AutoPage_|Optional| **Variant**|Valid only if  _SourceType_ is **xlConsolidation** . **True** to have Microsoft Excel create a page field for the consolidation. If _AutoPage_ is **False** , you must create the page field or fields.|
| _Reserved_|Optional| **Variant**|Not used by Microsoft Excel.|
| _BackgroundQuery_|Optional| **Variant**| **True** to have Excel perform queries for the report asynchronously (in the background). The default value is **False** .|
| _OptimizeCache_|Optional| **Variant**| **True** to optimize the PivotTable cache when it's constructed. The default value is **False** .|
| _PageFieldOrder_|Optional| **Variant**|The order in which page fields are added to the PivotTable report?s layout. Can be one of the following  **XlOrder** constants: **xlDownThenOver** or **xlOverThenDown** . The default value is **xlDownThenOver** .|
| _PageFieldWrapCount_|Optional| **Variant**|The number of page fields in each column or row in the PivotTable report. The default value is 0 (zero).|
| _ReadData_|Optional| **Variant**| **True** to create a PivotTable cache that contains all records from the external database; this cache can be very large. If _ReadData_ is **False** , you can set some of the fields asserver-based page fields before the data is actually read.|
| _Connection_|Optional| **Variant**|A string that contains ODBC settings that allow Excel to connect to an ODBC data source. The connection string has the form "ODBC;<connection string>". This argument overrides any previous setting for the  **[PivotCache](pivotcache-object-excel.md)** object?s **[Connection](pivotcache-connection-property-excel.md)** property.|

## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

