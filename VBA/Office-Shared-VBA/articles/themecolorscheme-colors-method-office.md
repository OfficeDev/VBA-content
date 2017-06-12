---
title: ThemeColorScheme.Colors Method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Colors
ms.assetid: 2ae73cd3-c1b7-1815-5b46-84c349c2535b
ms.date: 06/08/2017
---


# ThemeColorScheme.Colors Method (Office)

Gets an object that represents a color in the color scheme of a Microsoft Office theme.


## Syntax

 _expression_. **Colors**( **_Index_** )

 _expression_ An expression that returns a **ThemeColorScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**MsoThemeColorSchemeIndex**|The index value of the  **ThemeColor** object.|

### Return Value

ThemeColor


## Example

In the following example, the  **msoThemeAccent1** theme color is set to the color **Red** and then the scheme is saved to a file.


```
Dim tTheme As OfficeTheme 
Dim tcsThemeColorScheme As ThemeColorScheme 
Dim tcThemeColor As ThemeColor 
tcThemeColor.RGB = RGB(255, 0, 0) 
Set tcColorScheme.Colors(msoThemeAccent1) = tcThemeColor 
tcsThemeColorScheme.Save ("C:\myThemeColorScheme.xml") 

```


## See also


#### Concepts


[ThemeColorScheme Object](themecolorscheme-object-office.md)
#### Other resources


[ThemeColorScheme Object Members](themecolorscheme-members-office.md)

