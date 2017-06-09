---
title: ListRow Object (Excel)
keywords: vbaxl10.chm741072
f1_keywords:
- vbaxl10.chm741072
ms.prod: excel
api_name:
- Excel.ListRow
ms.assetid: ba3e4215-14b6-3dca-82d0-0951f9f2fc3e
ms.date: 06/08/2017
---


# ListRow Object (Excel)

Represents a row in a table. The  **ListRow** object is a member of the **[ListRows](listrows-object-excel.md)** collection.


## Remarks

The  **ListRows** collection contains all the rows in a list object.

Use the  **[ListRows](listobject-listrows-property-excel.md)** property of the **[ListObject](listobject-object-excel.md)** object to return a **ListRows** collection.


## Example

 The following example adds a new **ListRow** object to the default **ListObject** object in the first worksheet of the active workbook. Because no position is specified, a new row is added to the end of the table.


```
Dim wrksht As Worksheet 
Dim oListRow As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListRow = wrksht.ListObjects(1).ListRows.Add 

```


## Methods



|**Name**|
|:-----|
|[Delete](listrow-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](listrow-application-property-excel.md)|
|[Creator](listrow-creator-property-excel.md)|
|[Index](listrow-index-property-excel.md)|
|[Parent](listrow-parent-property-excel.md)|
|[Range](listrow-range-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
