---
title: ThemeFonts Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ThemeFonts
ms.assetid: 393865af-f008-d26c-5b82-9ae79766e511
---


# ThemeFonts Object (Office)

Represents a collection of major and minor fonts in the font scheme of a Microsoft Office theme.


## Example

The following example sets a  **ThemeFonts** object to a minor theme font.


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

