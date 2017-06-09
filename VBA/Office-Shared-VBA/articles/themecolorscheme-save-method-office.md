---
title: ThemeColorScheme.Save Method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.Save
ms.assetid: 5ca73773-583b-dbf4-6bde-bc6fa26c66a2
ms.date: 06/08/2017
---


# ThemeColorScheme.Save Method (Office)

Saves the color scheme of a Microsoft Office theme to a file.


## Syntax

 _expression_. **Save**( **_FileName_** )

 _expression_ An expression that returns a **ThemeColorScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file.|

## Example

The following example saves the color scheme for an Office theme to a file.


```
ThemeColorScheme.Save("C:\myThemeColorScheme.xml") 

```


## See also


#### Concepts


[ThemeColorScheme Object](themecolorscheme-object-office.md)
#### Other resources


[ThemeColorScheme Object Members](themecolorscheme-members-office.md)

