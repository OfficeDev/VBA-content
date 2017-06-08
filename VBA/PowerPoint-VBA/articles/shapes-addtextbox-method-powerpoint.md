---
title: Shapes.AddTextbox Method (PowerPoint)
keywords: vbapp10.chm543014
f1_keywords:
- vbapp10.chm543014
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddTextbox
ms.assetid: 0c7c6093-48f6-e1f1-1837-e69d6ef13e57
ms.date: 06/08/2017
---


# Shapes.AddTextbox Method (PowerPoint)

Creates a text box. Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the new text box.


## Syntax

 _expression_. **AddTextbox**( **_Orientation_**, **_Left_**, **_Top_**, **_Width_**, **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Orientation_|Required|**[MsoTextOrientation](http://msdn.microsoft.com/library/7e8d0e06-14dd-3ea1-a2e4-50375919517f%28Office.15%29.aspx)**|The text orientation. Some of these constants may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the text box relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the text box relative to the top edge of the slide.|
| _Width_|Required|**Single**|The width of the text box, measured in points.|
| _Height_|Required|**Single**| The height of the text box, measured in points.|

### Return Value

Shape


## Example

This example adds a text box that contains the text "Test Box" to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddTextbox(Type:=msoTextOrientationHorizontal, _ 
    Left:=100, Top:=100, Width:=200, Height:=50).TextFrame _ 
    .TextRange.Text = "Test Box"
```


## See also


#### Concepts


[Shapes Object](shapes-object-powerpoint.md)

