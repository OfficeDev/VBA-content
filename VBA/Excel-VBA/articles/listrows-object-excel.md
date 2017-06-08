---
title: ListRows Object (Excel)
keywords: vbaxl10.chm739072
f1_keywords:
- vbaxl10.chm739072
ms.prod: excel
api_name:
- Excel.ListRows
ms.assetid: e4035209-00a2-ea16-a3b9-2d23afe0b88a
ms.date: 06/08/2017
---


# ListRows Object (Excel)

A collection of all the  **[ListRow](listrow-object-excel.md)** objects in the specified **[ListObject](listobject-object-excel.md)** object.


## Remarks

 Each **ListRow** object represents a row in the table.


## Example

Use the  **[ListRows](listobject-listrows-property-excel.md)** property of the **[ListObject](listobject-object-excel.md)** object to return the **ListRows** collection. The following example adds a new row to the default **ListObject** object in the first worksheet of the workbook. Because no position is specified, a new row is added to the end of the table.


```
Set myNewRow = Worksheets(1).ListObject(0).ListRows.Add
```


## Methods



|**Name**|
|:-----|
|[Add](listrows-add-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](listrows-application-property-excel.md)|
|[Count](listrows-count-property-excel.md)|
|[Creator](listrows-creator-property-excel.md)|
|[Item](listrows-item-property-excel.md)|
|[Parent](listrows-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
