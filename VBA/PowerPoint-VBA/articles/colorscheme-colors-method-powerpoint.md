---
title: ColorScheme.Colors Method (PowerPoint)
keywords: vbapp10.chm537003
f1_keywords:
- vbapp10.chm537003
ms.prod: powerpoint
api_name:
- PowerPoint.ColorScheme.Colors
ms.assetid: ac910a40-9014-e709-491c-a8649fc08137
ms.date: 06/08/2017
---


# ColorScheme.Colors Method (PowerPoint)

Returns an  **[RGBColor](rgbcolor-object-powerpoint.md)** object that represents a single color in a color scheme.


## Syntax

 _expression_. **Colors**( **_SchemeColor_** )

 _expression_ A variable that represents a **ColorScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemeColor_|Required|**[PpColorSchemeIndex](ppcolorschemeindex-enumeration-powerpoint.md)**|The individual color in the specified color scheme.|

### Return Value

RGBColor


## Example

This example sets the title color for slides one and three in the active presentation.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```


## See also


#### Concepts


[ColorScheme Object](colorscheme-object-powerpoint.md)

