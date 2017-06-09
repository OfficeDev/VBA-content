---
title: ThemeEffectScheme.Load Method (Office)
ms.prod: office
api_name:
- Office.ThemeEffectScheme.Load
ms.assetid: 9bf428f7-bda8-c6d7-1688-05466f242280
ms.date: 06/08/2017
---


# ThemeEffectScheme.Load Method (Office)

Loads the effects scheme of a Microsoft Office theme from a file.


## Syntax

 _expression_. **Load**( **_FileName_** )

 _expression_ An expression that returns a **ThemeEffectScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the effect scheme file.|

## Example

The following example loads a theme effect scheme from a file.


```
tesEffectScheme.Load("C:\myThemeEffectScheme.eftx") 

```


## See also


#### Concepts


[ThemeEffectScheme Object](themeeffectscheme-object-office.md)
#### Other resources


[ThemeEffectScheme Object Members](themeeffectscheme-members-office.md)

