---
title: ListObject Object (Excel)
keywords: vbaxl10.chm733072
f1_keywords:
- vbaxl10.chm733072
ms.prod: excel
api_name:
- Excel.ListObject
ms.assetid: 46de6c4f-8ce0-0c7d-da59-6e52f5eab612
ms.date: 06/08/2017
---


# ListObject Object (Excel)

Represents a [ListObject Object (Excel)](listobject-object-excel.md) object in the **ListObjects** collection.


## Remarks

 The **ListObject** object is a member of the **[ListObjects](listobjects-object-excel.md)** collection. The **ListObjects** collection contains all the list objects on a worksheet.


## Example

Use the [ListObjects](worksheet-listobjects-property-excel.md) property of the **[Worksheet](worksheet-object-excel.md)** object to return a **ListObjects** collection. The following example adds a new **[ListRow](listrow-object-excel.md)** object to the default **ListObject** object in the first worksheet of the active workbook.


```
Dim wrksht As Worksheet 
Dim oListCol As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListCol = wrksht.ListObjects(1).ListRows.Add
```


## Methods



|**Name**|
|:-----|
|[Delete](listobject-delete-method-excel.md)|
|[ExportToVisio](listobject-exporttovisio-method-excel.md)|
|[Publish](listobject-publish-method-excel.md)|
|[Refresh](listobject-refresh-method-excel.md)|
|[Resize](listobject-resize-method-excel.md)|
|[Unlink](listobject-unlink-method-excel.md)|
|[Unlist](listobject-unlist-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Active](listobject-active-property-excel.md)|
|[AlternativeText](listobject-alternativetext-property-excel.md)|
|[Application](listobject-application-property-excel.md)|
|[AutoFilter](listobject-autofilter-property-excel.md)|
|[Comment](listobject-comment-property-excel.md)|
|[Creator](listobject-creator-property-excel.md)|
|[DataBodyRange](listobject-databodyrange-property-excel.md)|
|[DisplayName](listobject-displayname-property-excel.md)|
|[DisplayRightToLeft](listobject-displayrighttoleft-property-excel.md)|
|[HeaderRowRange](listobject-headerrowrange-property-excel.md)|
|[InsertRowRange](listobject-insertrowrange-property-excel.md)|
|[ListColumns](listobject-listcolumns-property-excel.md)|
|[ListRows](listobject-listrows-property-excel.md)|
|[Name](listobject-name-property-excel.md)|
|[Parent](listobject-parent-property-excel.md)|
|[QueryTable](listobject-querytable-property-excel.md)|
|[Range](listobject-range-property-excel.md)|
|[SharePointURL](listobject-sharepointurl-property-excel.md)|
|[ShowAutoFilter](listobject-showautofilter-property-excel.md)|
|[ShowAutoFilterDropDown](listobject-showautofilterdropdown-property-excel.md)|
|[ShowHeaders](listobject-showheaders-property-excel.md)|
|[ShowTableStyleColumnStripes](listobject-showtablestylecolumnstripes-property-excel.md)|
|[ShowTableStyleFirstColumn](listobject-showtablestylefirstcolumn-property-excel.md)|
|[ShowTableStyleLastColumn](listobject-showtablestylelastcolumn-property-excel.md)|
|[ShowTableStyleRowStripes](listobject-showtablestylerowstripes-property-excel.md)|
|[ShowTotals](listobject-showtotals-property-excel.md)|
|[Slicers](listobject-slicers-property-excel.md)|
|[Sort](listobject-sort-property-excel.md)|
|[SourceType](listobject-sourcetype-property-excel.md)|
|[Summary](listobject-summary-property-excel.md)|
|[TableObject](listobject-tableobject-property-excel.md)|
|[TableStyle](listobject-tablestyle-property-excel.md)|
|[TotalsRowRange](listobject-totalsrowrange-property-excel.md)|
|[XmlMap](listobject-xmlmap-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
