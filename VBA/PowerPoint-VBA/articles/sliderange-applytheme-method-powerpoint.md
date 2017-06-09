---
title: SlideRange.ApplyTheme Method (PowerPoint)
keywords: vbapp10.chm532039
f1_keywords:
- vbapp10.chm532039
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.ApplyTheme
ms.assetid: 779ca8d3-e235-7f65-1a2f-b5233517da1f
ms.date: 06/08/2017
---


# SlideRange.ApplyTheme Method (PowerPoint)

Applies a theme or design template to the specified range of slides.


## Syntax

 _expression_. **ApplyTheme**( **_themeName_** )

 _expression_ A variable that represents a **SlideRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _themeName_|Required|**String**|The path and name of the theme file (.thmx) or design template file (.pot) to apply to the  **SlideRange** object.|

## Example

This example applies a saved theme to the specified range of slides:


```vb
ActivePresentation.Slides.Range(Array(1, 3)).ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.thmx"
```

This example applies a saved design template to the specified range of slides:




```vb
ActivePresentation.Slides.Range(Array(1, 3)).ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.pot"
```


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

