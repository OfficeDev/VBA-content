---
title: SparklineGroups Object (Excel)
keywords: vbaxl10.chm868072
f1_keywords:
- vbaxl10.chm868072
ms.prod: excel
api_name:
- Excel.SparklineGroups
ms.assetid: 9bc6be34-fa2e-8652-ca92-fa9630b4d7a6
ms.date: 06/08/2017
---


# SparklineGroups Object (Excel)

Represents a collection of sparkline groups.


## Remarks

The  **SparklineGroups** object can contain multiple **[SparklineGroup](sparklinegroup-object-excel.md)** objects.

Use the  **[SparklineGroups](range-sparklinegroups-property-excel.md)** property of the **[Range](range-object-excel.md)** object to return an existing **SparklineGroups** collection from its parent range.

Use the  **[Add](sparklinegroups-add-method-excel.md)** method to create a group of new sparklines.

Use the  **[Group](sparklinegroups-group-method-excel.md)** method to create a group of existing sparklines.


## Example

This example selects the range A1:A4 and groups the sparklines in that range. If the sparklines in the sparkline group are line sparklines, the markers are displayed in red.


```vb
Range("A1:A4").Select 
Selection.SparklineGroups.Group Location := Range("A1") 
Selection.SparklineGroups.Item(1).Points.Markers.Visible = True 
Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 255
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


