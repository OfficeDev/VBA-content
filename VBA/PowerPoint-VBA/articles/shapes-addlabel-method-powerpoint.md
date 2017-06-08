---
title: Shapes.AddLabel Method (PowerPoint)
keywords: vbapp10.chm543008
f1_keywords:
- vbapp10.chm543008
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddLabel
ms.assetid: b744daf1-5b99-9649-8b97-d3f2193373c1
ms.date: 06/08/2017
---


# Shapes.AddLabel Method (PowerPoint)

Creates a label. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new label.


## Syntax

 _expression_. **AddLabel**( **_Orientation_**, **_Left_**, **_Top_**, **_Width_**, **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required|**[MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The text orientation. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the label relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the label relative to the top edge of the slide.|
| _Width_|Required|**Single**|The width of the label, measured in points.|
| _Height_|Required|**Single**|The height of the label, measured in points.|

### Return Value

Shape


## Example

This example adds a vertical label that contains the text "Test Label" to myDocument.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddLabel(Orientation:=msoTextOrientationVerticalFarEast, _ 
    Left:=100, Top:=100, Width:=60, Height:=150).TextFrame _ 
    .TextRange.Text = "Test Label"
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

