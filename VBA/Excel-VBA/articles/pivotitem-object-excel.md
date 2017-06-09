---
title: PivotItem Object (Excel)
keywords: vbaxl10.chm245072
f1_keywords:
- vbaxl10.chm245072
ms.prod: excel
api_name:
- Excel.PivotItem
ms.assetid: 5829a1d9-0924-9ce8-1120-229e4595285a
ms.date: 06/08/2017
---


# PivotItem Object (Excel)

Represents an item in a PivotTable field.


## Remarks

 The items are the individual data entries in a field category. The **PivotItem** object is a member of the **[PivotItems](pivotitems-object-excel.md)** collection. The **PivotItems** collection contains all the items in a **PivotField** object.


## Example

Use  **[PivotItems](pivotfield-pivotitems-method-excel.md)** ( _index_ ), where _index_ is the item index number or name, to return a single **PivotItem** object. The following example hides all entries in the first PivotTable report on Sheet3 that contain "1998" in the Year field.


```
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").PivotItems("1998").Visible = False
```


## Methods



|**Name**|
|:-----|
|[Delete](pivotitem-delete-method-excel.md)|
|[DrillTo](pivotitem-drillto-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pivotitem-application-property-excel.md)|
|[Caption](pivotitem-caption-property-excel.md)|
|[ChildItems](pivotitem-childitems-property-excel.md)|
|[Creator](pivotitem-creator-property-excel.md)|
|[DataRange](pivotitem-datarange-property-excel.md)|
|[DrilledDown](pivotitem-drilleddown-property-excel.md)|
|[Formula](pivotitem-formula-property-excel.md)|
|[IsCalculated](pivotitem-iscalculated-property-excel.md)|
|[LabelRange](pivotitem-labelrange-property-excel.md)|
|[Name](pivotitem-name-property-excel.md)|
|[Parent](pivotitem-parent-property-excel.md)|
|[ParentItem](pivotitem-parentitem-property-excel.md)|
|[ParentShowDetail](pivotitem-parentshowdetail-property-excel.md)|
|[Position](pivotitem-position-property-excel.md)|
|[RecordCount](pivotitem-recordcount-property-excel.md)|
|[ShowDetail](pivotitem-showdetail-property-excel.md)|
|[SourceName](pivotitem-sourcename-property-excel.md)|
|[SourceNameStandard](pivotitem-sourcenamestandard-property-excel.md)|
|[StandardFormula](pivotitem-standardformula-property-excel.md)|
|[Value](pivotitem-value-property-excel.md)|
|[Visible](pivotitem-visible-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
