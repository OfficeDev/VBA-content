---
title: Master.ApplyTheme Method (PowerPoint)
keywords: vbapp10.chm533019
f1_keywords:
- vbapp10.chm533019
ms.prod: powerpoint
api_name:
- PowerPoint.Master.ApplyTheme
ms.assetid: ae30318b-20e6-4eae-df4c-1f159fd77d6a
ms.date: 06/08/2017
---


# Master.ApplyTheme Method (PowerPoint)

Applies a theme or design template to the specified slide master, title master, handout master, notes master, or design master.


## Syntax

 _expression_. **ApplyTheme**( **_themeName_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _themeName_|Required|**String**|The path and name of the theme file (.thmx) or design template file (.pot) to apply to the  **Master** object.|

## Example

This example applies a saved theme to a slide master:


```vb
ActivePresentation.SlideMaster.ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.thmx"
```

This example applies a saved design template to a slide master:




```vb
ActivePresentation.SlideMaster.ApplyTheme "C:\Program Files\Microsoft Office\Templates\MyTheme.pot"
```


## See also


#### Concepts


[Master Object](master-object-powerpoint.md)

