---
title: PivotItem.Name Property (Excel)
keywords: vbaxl10.chm246078
f1_keywords:
- vbaxl10.chm246078
ms.prod: excel
api_name:
- Excel.PivotItem.Name
ms.assetid: b3861675-1f05-9e0d-442c-1cd95385ca09
ms.date: 06/08/2017
---


# PivotItem.Name Property (Excel)

Returns or sets a  **String** value representing the name of the object.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **PivotItem** object.


## Remarks

The following table shows example values of the  **Name** property and related properties given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



|**Property**|**Value (OLAP data source)**|**Value (non-OLAP data source)**|
|:-----|:-----|:-----|
| **[Caption](pivotitem-caption-property-excel.md)**|Paris|Paris|
| **[Name](pivotitem-name-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **[SourceName](pivotitem-sourcename-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|(Same as SQL property value, read-only)|
| **[Value](pivotitem-value-property-excel.md)**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
When specifying an index into the  **[PivotItems](pivotitems-object-excel.md)** collection, you can use the syntax shown in the following table.



|**Syntax (OLAP data source)**|**Syntax (non-OLAP data source)**|
|:-----|:-----|
|expression.PivotItems("[Europe].[France].[Paris]")|expression.PivotItems("Paris")|
When using the  **[Item](iconcriteria-item-property-excel.md)** property to reference a specific member of a collection, you can use the text index name as shown in the following table.



|**Name (OLAP data source)**|**Name (non-OLAP data source)**|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|

## See also


#### Concepts


[PivotItem Object](pivotitem-object-excel.md)

