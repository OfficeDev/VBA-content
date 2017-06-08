---
title: RGBColor.RGB Property (PowerPoint)
keywords: vbapp10.chm538003
f1_keywords:
- vbapp10.chm538003
ms.prod: powerpoint
api_name:
- PowerPoint.RGBColor.RGB
ms.assetid: 0535b619-1d3d-a106-8b99-46ea5c02917f
ms.date: 06/08/2017
---


# RGBColor.RGB Property (PowerPoint)

Returns or sets the red-green-blue (RGB) value of a specified color-scheme color or extra color when used with a  **PpColorSchemeIndex** constant. Read/write.


## Syntax

 _expression_. **RGB**

 _expression_ A variable that represents a **RGBColor** object.


### Return Value

MsoRGBType


## Remarks

Use the  **Colors** method to return a **RGBColor** object.

The value of the  **RGB** property can be one of these **PpColorSchemeIndex** constants.


||
|:-----|
|**ppAccent1**|
|**ppAccent2**|
|**ppAccent3**|
|**ppBackground**|
|**ppFill**|
|**ppForeground**|
|**ppShadow**|
|**ppTitle**|

## Example

This example displays the value of the red, green, and blue components of the fill forecolor for shape one on slide one in the active document.


```vb
Set myDocument = ActivePresentation.Slides(1)

c = myDocument.Shapes(1).Fill.ForeColor.RGB

redComponent = c Mod 256

greenComponent = c \ 256 Mod 256

blueComponent = c \ 65536 Mod 256

MsgBox "RGB components: " &; redComponent &; _
    ", " &; greenComponent &; ", " &; blueComponent
```


## See also


#### Concepts


[RGBColor Object](rgbcolor-object-powerpoint.md)

