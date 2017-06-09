---
title: PivotTable.DisplayNullString Property (Excel)
keywords: vbaxl10.chm235105
f1_keywords:
- vbaxl10.chm235105
ms.prod: excel
api_name:
- Excel.PivotTable.DisplayNullString
ms.assetid: ad2ce480-9fc9-d069-5526-4f819e236967
ms.date: 06/08/2017
---


# PivotTable.DisplayNullString Property (Excel)

 **True** if the PivotTable report displays a custom string in cells that contain null values. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **DisplayNullString**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Use the  **[NullString](pivottable-displaynullstring-property-excel.md)** property to set the custom null string.


## Example

This example causes the PivotTable report to display "NA" in cells that contain null values.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .NullString = "NA" 
 .DisplayNullString = True 
End With
```

This example causes the PivotTable report to display 0 (zero) in cells that contain null values.




```vb
Worksheets(1).PivotTables("Pivot1").DisplayNullString = False
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

