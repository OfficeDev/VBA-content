---
title: Slide.ApplyTheme Method (PowerPoint)
keywords: vbapp10.chm531035
f1_keywords:
- vbapp10.chm531035
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.ApplyTheme
ms.assetid: 70fff6cd-0541-dff8-754e-e8ee1a46dc2b
ms.date: 06/08/2017
---


# Slide.ApplyTheme Method (PowerPoint)

Applies a theme or design template to the specified slide.


## Syntax

 _expression_. **ApplyTheme**( **_themeName_** )

 _expression_ A variable that represents a **Slide** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _themeName_|Required|**String**|The path and name of the theme file (.thmx) or design template file (.pot) to apply to the  **Slide** object.|

## Example

This example applies a saved theme to the specified slide:


```vb
ActivePresentation.Slides(1).ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.thmx"
```

This example applies a saved design template to the specified slide:




```vb
ActivePresentation.Slides(1).ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.pot"
```


## See also


#### Concepts


[Slide Object](slide-object-powerpoint.md)

