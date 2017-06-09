---
title: ListColumn Object (Excel)
keywords: vbaxl10.chm737072
f1_keywords:
- vbaxl10.chm737072
ms.prod: excel
api_name:
- Excel.ListColumn
ms.assetid: c2060e4a-2340-c606-f272-1e4dad6964d0
ms.date: 06/08/2017
---


# ListColumn Object (Excel)

Represents a column in a table.


## Remarks

 The **ListColumn** object is a member of the **[ListColumns](listcolumns-object-excel.md)** collection. The **ListColumns** collection contains all the columns in a table ( **[ListObject](listobject-object-excel.md)** object).

Use the [ListColumns](listobject-listcolumns-property-excel.md) property of the **ListObject** object to return a **[ListColumns](listcolumns-object-excel.md)** collection.


## Example

The following example adds a new  **ListColumn** object to the default **ListObject** object in the first worksheet of the active workbook. Because no position is specified, a new rightmost column is added.


```
Sub AddListColumn() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns.Add 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Delete](listcolumn-delete-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](listcolumn-application-property-excel.md)|
|[Creator](listcolumn-creator-property-excel.md)|
|[DataBodyRange](listcolumn-databodyrange-property-excel.md)|
|[Index](listcolumn-index-property-excel.md)|
|[Name](listcolumn-name-property-excel.md)|
|[Parent](listcolumn-parent-property-excel.md)|
|[Range](listcolumn-range-property-excel.md)|
|[Total](listcolumn-total-property-excel.md)|
|[TotalsCalculation](listcolumn-totalscalculation-property-excel.md)|
|[XPath](listcolumn-xpath-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
