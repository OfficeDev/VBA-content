---
title: ThemeFontScheme.Load Method (Office)
ms.prod: office
api_name:
- Office.ThemeFontScheme.Load
ms.assetid: a9ac928e-904f-70bd-1e96-932243204d73
ms.date: 06/08/2017
---


# ThemeFontScheme.Load Method (Office)

Loads the font scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. **Load**( **_FileName_** )

 _expression_ An expression that returns a **ThemeFontScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the font scheme file.|

## Example

The following example loads a theme font scheme from a file.


```
ThemeFontScheme.Load ("C:\myThemeFontScheme.xml")
```


## See also


#### Concepts


[ThemeFontScheme Object](themefontscheme-object-office.md)
#### Other resources


[ThemeFontScheme Object Members](themefontscheme-members-office.md)

