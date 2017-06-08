---
title: ThemeColorScheme.Load Method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Load
ms.assetid: 636f14c1-4178-ef12-e22b-4d948719cced
ms.date: 06/08/2017
---


# ThemeColorScheme.Load Method (Office)

Loads the color scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. **Load**( **_FileName_** )

 _expression_ An expression that returns a **ThemeColorScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the color theme file.|

## Example

The following example loads a theme color scheme from a file.


```
ThemeColorScheme.Load ("C:\myThemeColorScheme.xml") 

```


## See also


#### Concepts


[ThemeColorScheme Object](themecolorscheme-object-office.md)
#### Other resources


[ThemeColorScheme Object Members](themecolorscheme-members-office.md)

