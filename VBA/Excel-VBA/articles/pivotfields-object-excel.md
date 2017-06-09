---
title: PivotFields Object (Excel)
keywords: vbaxl10.chm241072
f1_keywords:
- vbaxl10.chm241072
ms.prod: excel
api_name:
- Excel.PivotFields
ms.assetid: 018d4cea-09ea-d4be-baef-5fd55062935b
ms.date: 06/08/2017
---


# PivotFields Object (Excel)

A collection of all the  **[PivotField](pivotfield-object-excel.md)** objects in a PivotTable report.


## Remarks

In some cases, it may be easier to use one of the properties that returns a subset of the PivotTable fields. The following accessor methods are available:


-  **[ColumnFields](pivottable-columnfields-property-excel.md)** property
    
-  **[DataFields](pivottable-datafields-property-excel.md)** property
    
-  **[HiddenFields](pivottable-hiddenfields-property-excel.md)** property
    
-  **[PageFields](pivottable-pagefields-property-excel.md)** property
    
-  **[RowFields](pivottable-rowfields-property-excel.md)** property
    
-  **[VisibleFields](pivottable-visiblefields-property-excel.md)** property
    

## Example

Use the  **PivotFields** method of the **PivotTable** object to return the **PivotFields** collection. The following example enumerates the field names in the first PivotTable report on Sheet3.


```
With Worksheets("sheet3").PivotTables(1) 
 For i = 1 To .PivotFields.Count 
 MsgBox .PivotFields(i).Name 
 Next 
End With
```

Use  **[PivotFields](pivottable-pivotfields-method-excel.md)** ( _index_ ), where _index_ is the field name or index number, to return a single **PivotField** object. The following example makes the Year field a row field in the first PivotTable report on Sheet3.




```
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").Orientation = xlRowField
```


## Methods



|**Name**|
|:-----|
|[Item](pivotfields-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pivotfields-application-property-excel.md)|
|[Count](pivotfields-count-property-excel.md)|
|[Creator](pivotfields-creator-property-excel.md)|
|[Parent](pivotfields-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
