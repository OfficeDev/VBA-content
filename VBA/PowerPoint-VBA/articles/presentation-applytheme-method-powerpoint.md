---
title: Presentation.ApplyTheme Method (PowerPoint)
keywords: vbapp10.chm583105
f1_keywords:
- vbapp10.chm583105
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ApplyTheme
ms.assetid: e403614b-fc39-98e0-e707-501394aacfa1
ms.date: 06/08/2017
---


# Presentation.ApplyTheme Method (PowerPoint)

Applies a theme or design template to the specified presentation.


## Syntax

 _expression_. **ApplyTheme**( **_themeName_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _themeName_|Required|**String**|The path and name of the theme file (.thmx) or design template file (.pot) to apply to the  **Presentation** object.|

## Example

This example applies a saved theme to the active presentation:


```vb
ActivePresentation.ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.thmx"
```

This example applies a saved design template to the active presentation:




```vb
ActivePresentation.ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.pot"
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

