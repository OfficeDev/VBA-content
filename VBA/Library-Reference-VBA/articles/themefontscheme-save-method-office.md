---
title: ThemeFontScheme.Save Method (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ThemeFontScheme.Save
ms.assetid: 4adbeac7-b5cf-327e-f999-4dd2d721755d
---


# ThemeFontScheme.Save Method (Office)

Saves the font scheme of a Microsoft Office theme to a file.


## Syntax

 _expression_. **Save**( ** _FileName_** )

 _expression_ An expression that returns a **ThemeFontScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file.|

## Example

The following example saves the Office theme font scheme to a file. 


```
ThemeFontScheme.Save("C:\myThemeFontScheme.xml")
```


## See also


#### Concepts


[ThemeFontScheme Object](themefontscheme-object-office.md)

