---
title: PivotItems Object (Excel)
keywords: vbaxl10.chm247072
f1_keywords:
- vbaxl10.chm247072
ms.prod: excel
api_name:
- Excel.PivotItems
ms.assetid: df47021a-2b06-fa10-5712-58956c7ffe07
ms.date: 06/08/2017
---


# PivotItems Object (Excel)

A collection of all the  **[PivotItem](pivotitem-object-excel.md)** objects in a PivotTable field.


## Remarks

 The items are the individual data entries in a field category.


## Example

Use the  **[PivotItems](pivotfield-pivotitems-method-excel.md)** method to return the **[PivotItems](pivotitems-object-excel.md)** collection. The following example creates an enumerated list of field names and the items contained in those fields for the first PivotTable report on Sheet4.


```
Worksheets("sheet4").Activate 
With Worksheets("sheet3").PivotTables(1) 
 c = 1 
 For i = 1 To .PivotFields.Count 
 r = 1 
 Cells(r, c) = .PivotFields(i).Name 
 r = r + 1 
 For x = 1 To .PivotFields(i).PivotItems.Count 
 Cells(r, c) = .PivotFields(i).PivotItems(x).Name 
 r = r + 1 
 Next 
 c = c + 1 
 Next 
End With
```

Use  **PivotItems** ( _index_ ), where _index_ is the item index number or name, to return a single **PivotItem** object. The following example hides all entries in the first PivotTable report on Sheet3 that contain "1998" in the Year field.




```
Worksheets("sheet3").PivotTables(1) _ 
 .PivotFields("year").PivotItems("1998").Visible = False
```


## Methods



|**Name**|
|:-----|
|[Add](pivotitems-add-method-excel.md)|
|[Item](pivotitems-item-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](pivotitems-application-property-excel.md)|
|[Count](pivotitems-count-property-excel.md)|
|[Creator](pivotitems-creator-property-excel.md)|
|[Parent](pivotitems-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
