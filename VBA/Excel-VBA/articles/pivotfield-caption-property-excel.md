---
title: PivotField.Caption Property (Excel)
keywords: vbaxl10.chm240124
f1_keywords:
- vbaxl10.chm240124
ms.prod: excel
api_name:
- Excel.PivotField.Caption
ms.assetid: 7cd928bf-3f69-0950-5b51-9168192c349e
ms.date: 06/08/2017
---


# PivotField.Caption Property (Excel)

Returns a  **String** value that represents the label text for the pivot field.


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

The following table shows example values of the  **Caption** property and related properties, given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



| <strong>Property</strong>                                                                                                                                 | <strong>Value (OLAP data source)</strong>    | <strong>Value (non-OLAP data source)</strong> |
|:----------------------------------------------------------------------------------------------------------------------------------------------------------|:---------------------------------------------|:----------------------------------------------|
| <strong><a href="pivotfield-caption-property-excel.md" data-raw-source="[Caption](pivotfield-caption-property-excel.md)">Caption</a></strong>             | Paris                                        | Paris                                         |
| <strong><a href="pivotfield-name-property-excel.md" data-raw-source="[Name](pivotfield-name-property-excel.md)">Name</a></strong>                         | [Europe].[France].[Paris] &nbsp; (read-only) | Paris                                         |
| <strong><a href="pivotfield-sourcename-property-excel.md" data-raw-source="[SourceName](pivotfield-sourcename-property-excel.md)">SourceName</a></strong> | [Europe].[France].[Paris] &nbsp;(read-only)  | (Same as the SQL property value; read-only)   |
| <strong><a href="pivotfield-value-property-excel.md" data-raw-source="[Value](pivotfield-value-property-excel.md)">Value</a></strong>                     | [Europe].[France].[Paris]  &nbsp;(read-only) | Paris                                         |

When specifying an index into the  **[PivotItems](pivotitems-object-excel.md)** collection, you can use the syntax shown in the following table.



| <strong>Syntax (OLAP data source)</strong>         | <strong>Syntax (non-OLAP data source)</strong> |
|:---------------------------------------------------|:-----------------------------------------------|
| expression.PivotItems("[Europe].[France].[Paris]") | expression.PivotItems("Paris")                 |

When using the  **Item** property to reference a specific member of a collection, you can use the text index names shown in the following table.



|**Name (OLAP data source)**|**Name (non-OLAP data source)**|
|:-----|:-----|
|[Europe].[France].[Paris]|Paris|

## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

