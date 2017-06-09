---
title: ListObjects Object (Excel)
keywords: vbaxl10.chm731072
f1_keywords:
- vbaxl10.chm731072
ms.prod: excel
api_name:
- Excel.ListObjects
ms.assetid: 3a888055-1ed0-d37d-0586-ced999dc1c42
ms.date: 06/08/2017
---


# ListObjects Object (Excel)

A collection of all the  **[ListObject](listobject-object-excel.md)** objects on a worksheet. Each **ListObject** object represents a table in the worksheet.


## Remarks

Use the  **[ListObjects](worksheet-listobjects-property-excel.md)** property of the[Worksheet](worksheet-object-excel.md) object to return the **ListObjects** collection.


## Example

 The following example creates a new **ListObjects** collection which represents all the tables in a worksheet.


```
Set myWorksheetLists = Worksheets(1).ListObjects
```


## Methods



|**Name**|
|:-----|
|[Add](listobjects-add-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](listobjects-application-property-excel.md)|
|[Count](listobjects-count-property-excel.md)|
|[Creator](listobjects-creator-property-excel.md)|
|[Item](listobjects-item-property-excel.md)|
|[Parent](listobjects-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
