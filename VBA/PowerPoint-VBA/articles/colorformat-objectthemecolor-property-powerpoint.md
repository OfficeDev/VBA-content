---
title: ColorFormat.ObjectThemeColor Property (PowerPoint)
keywords: vbapp10.chm506006
f1_keywords:
- vbapp10.chm506006
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.ObjectThemeColor
ms.assetid: 40264b94-b16d-2a52-9adc-8e8510ec581d
ms.date: 06/08/2017
---


# ColorFormat.ObjectThemeColor Property (PowerPoint)

Returns or sets the theme color of the specified  **ColorFormat** object. Read/Write.


## Syntax

 _expression_. **ObjectThemeColor**

 _expression_ An expression that returns a **ColorFormat** object.


### Return Value

MsoThemeColorIndex


## Remarks

The value of the  **ObjectThemeColor** property can be one of these **[MsoThemeColorIndex](http://msdn.microsoft.com/library/2281eafa-c8f0-d620-d0eb-c301dfb6a426%28Office.15%29.aspx)** constants.


## Example

The following example shows how to use the  **ObjectThemeColor** property to get the theme color of the foreground fill of shape one on slide one of the active presentation.


```vb
Public Sub ObjectThemeColor_Example() 
 
    Debug.Print ActivePresentation.Slides(1).Shapes(1).Fill.ForeColor.ObjectThemeColor 
     
End Sub
```


## See also


#### Concepts


[ColorFormat Object](colorformat-object-powerpoint.md)

