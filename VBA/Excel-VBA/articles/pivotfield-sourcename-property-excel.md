---
title: PivotField.SourceName Property (Excel)
keywords: vbaxl10.chm240093
f1_keywords:
- vbaxl10.chm240093
ms.prod: excel
api_name:
- Excel.PivotField.SourceName
ms.assetid: d18eb5a0-d44c-9f04-45b1-94cdf468c13e
ms.date: 06/08/2017
---


# PivotField.SourceName Property (Excel)

Returns a  **String** value that represents the specified object?s name as it appears in the original source data for the specified PivotTable report.


## Syntax

 _expression_ . **SourceName**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The value of this property might be different from the current item name if the user renamed the item after creating the PivotTable report.

The following table shows example values of the  **SourceName** property and related properties, given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



|**Property**|**Value (OLAP data source)**|**Value (non-OLAP data source)**|
|:-----|:-----|:-----|
| **[Caption](pivotfield-caption-property-excel.md)**|Paris|Paris|
| **[Name](pivotfield-name-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **[SourceName](pivotfield-sourcename-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|(same as SQL property value, read-only)|
| **[Value](pivotfield-value-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
When specifying an index into the  **[PivotItems](pivotitems-object-excel.md)** collection, you can use the syntax shown in the following table.



|**Syntax (OLAP data source)**|**Syntax (non-OLAP data source)**|
|:-----|:-----|
|expression.PivotItems("[Europe].[France].[Paris]")|expression.PivotItems("Paris")|
When using the  **[Item](pivotitems-item-method-excel.md)** property to reference a specific member of a collection, you can use the text index names, as shown in the following table.



|**Name (OLAP data source)**|**Name (non-OLAP data source)**|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|

## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

