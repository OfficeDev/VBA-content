---
title: ThreeDFormat.SetThreeDFormat Method (Excel)
keywords: vbaxl10.chm119003
f1_keywords:
- vbaxl10.chm119003
ms.prod: excel
api_name:
- Excel.ThreeDFormat.SetThreeDFormat
ms.assetid: 64315607-991a-426e-e931-78432558832e
ms.date: 06/08/2017
---


# ThreeDFormat.SetThreeDFormat Method (Excel)

Sets the preset extrusion format. Each preset extrusion format contains a set of preset values for the various properties of the extrusion.


## Syntax

 _expression_ . **SetThreeDFormat**( **_PresetThreeDFormat_** )

 _expression_ A variable that represents a **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetThreeDFormat_|Required| **[MsoPresetThreeDFormat](http://msdn.microsoft.com/library/9d362115-1979-d079-d7e5-2e7788da614b%28Office.15%29.aspx)**|Specifies a preset extrusion format that corresponds to one of the options (numbered from left to right, from top to bottom) displayed when you click the  **3-D** button on the **Drawing** toolbar.|

## Remarks

This method sets the  **[PresetThreeDFormat](threedformat-presetthreedformat-property-excel.md)** property to the format specified by the _PresetThreeDFormat_ argument.


## Example

This example adds an oval to  `myDocument` and sets its extrusion format to 3D Style 12.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, _ 
        30, 30, 50, 25).ThreeD 
    .Visible = True 
    .SetThreeDFormat msoThreeD12 
End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

