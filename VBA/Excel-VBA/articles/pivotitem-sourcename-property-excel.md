---
title: PivotItem.SourceName Property (Excel)
keywords: vbaxl10.chm246083
f1_keywords:
- vbaxl10.chm246083
ms.prod: excel
api_name:
- Excel.PivotItem.SourceName
ms.assetid: 9222dcaf-fb60-45c1-a230-4eb7201e1c2a
ms.date: 06/08/2017
---


# PivotItem.SourceName Property (Excel)

Returns a  **Variant** value that represents the specified object?s name as it appears in the original source data for the specified PivotTable report.


## Syntax

 _expression_ . **SourceName**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

The value of this property might be different from the current item name if the user renamed the item after creating the PivotTable report.

The following table shows example values of the  **SourceName** property and related properties, given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



|**Property**|**Value (OLAP data source)**|**Value (non-OLAP data source)**|
|:-----|:-----|:-----|
| **[Caption](pivotitem-caption-property-excel.md)**|Paris|Paris|
| **[Name](pivotitem-name-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **[SourceName](pivotitem-sourcename-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|(same as SQL property value, read-only)|
| **[Value](pivotitem-value-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
When specifying an index into the  **[PivotItems](pivotitems-object-excel.md)** collection, you can use the syntax shown in the following table.



|**Syntax (OLAP data source)**|**Syntax (non-OLAP data source)**|
|:-----|:-----|
|expression.PivotItems("[Europe].[France].[Paris]")|expression.PivotItems("Paris")|
When using the  **[Item](pivotitems-item-method-excel.md)** property to reference a specific member of a collection, you can use the text index names, as shown in the following table.



|**Name (OLAP data source)**|**Name (non-OLAP data source)**|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|

## Example

This example displays the original name (the name from the source database) of the item that contains the active cell.


```vb
Worksheets("Sheet1").Activate 
ActiveSheet.PivotTables(1).PivotSelect "1998", xlDataAndLabel 
MsgBox "The original item name is " &; _ 
 ActiveCell.PivotItem.SourceName
```


## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

