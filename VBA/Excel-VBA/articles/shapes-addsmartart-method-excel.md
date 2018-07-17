---
title: Shapes.AddSmartArt Method (Excel)
keywords: vbaxl10.chm638095
f1_keywords:
- vbaxl10.chm638095
ms.prod: excel
api_name:
- Excel.Shapes.AddSmartArt
ms.assetid: e18a53ef-7649-34be-a264-aa545dd3d012
ms.date: 06/08/2017
---


# Shapes.AddSmartArt Method (Excel)

Creates a new SmartArt graphic with the specified layout. 


## Syntax

 _expression_ . **AddSmartArt**( **_Layout_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **[Shapes](shapes-object-excel.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **[SmartArtLayout](http://msdn.microsoft.com/library/f8d9db83-86f7-4830-096d-5d15368ab6b1%28Office.15%29.aspx)**|An object that represents the layout to use.|
| _Left_|Optional| **Variant**|The distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
| _Top_|Optional| **Variant**|The distance, in points, from the top edge of the object to the top edge of the worksheet.|
| _Width_|Optional| **Variant**|The width, in points, of the object.|
| _Height_|Optional| **Variant**|The width, in points, of the object.|

### Return Value

Shape


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

