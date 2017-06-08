---
title: PivotTables Object (Excel)
keywords: vbaxl10.chm237072
f1_keywords:
- vbaxl10.chm237072
ms.prod: excel
api_name:
- Excel.PivotTables
ms.assetid: 5beb33ac-a0fb-3f78-8fdc-d05719512214
ms.date: 06/08/2017
---


# PivotTables Object (Excel)

A collection of all the  **[PivotTable](pivottable-object-excel.md)** objects in the specified workbook.


## Remarks


 **Note**  The [Workbook.PivotTables](workbook-pivottables-property-excel.md) property (which is new for Office) does not return all the **PivotTable** objects in the workbook; instead it returns only those associated with decoupled PivotCharts. However,[Worksheet.PivotTables](worksheet-pivottables-method-excel.md) returns all the **PivotTable** objects in the worksheet, irrespective of whether they are associated with decoupled PivotCharts.

Because PivotTable report programming can be complex, it's generally easiest to record PivotTable report actions and then revise the recorded code.


## Example

Use the  **[PivotTables](worksheet-pivottables-method-excel.md)** method to return the **PivotTables** collection. The following example displays the number of PivotTable reports on Sheet3.


```
MsgBox Worksheets("sheet3").PivotTables.Count
```

Use the  **[PivotTableWizard](worksheet-pivottablewizard-method-excel.md)** method to create a new PivotTable report and add it to the collection. The following example creates a new PivotTable report from a Microsoft Excel database (contained in the range A1:C100).




```
ActiveSheet.PivotTableWizard xlDatabase, Range("A1:C100")
```

Use  **PivotTables** ( _index_ ), where _index_ is the PivotTable index number or name, to return a single **PivotTable** object. The following example makes the Year field a row field in the first PivotTable report on Sheet3.




```
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods



|**Name**|
|:-----|
|[Add](pivottables-add-method-excel.md)|
|[Item](pivottables-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pivottables-application-property-excel.md)|
|[Count](pivottables-count-property-excel.md)|
|[Creator](pivottables-creator-property-excel.md)|
|[Parent](pivottables-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
