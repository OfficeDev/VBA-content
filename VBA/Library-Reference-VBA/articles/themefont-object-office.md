---
title: ThemeFont Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.ThemeFont
ms.assetid: 1a9f1365-c392-3d04-74db-333ac111114a
---


# ThemeFont Object (Office)

Represents a container for the font schemes of a Microsoft Office theme.


## Example

The following example sets the Headings font scheme in a Microsoft Office theme to a Latin scheme. 


```vb
Dim tTheme As OfficeTheme 
Dim tfThemeFontScheme As ThemeFontScheme 
Dim tfThemeFont As ThemeFont 
Set tfThemeFontScheme = tTheme.ThemeFontScheme 
Set tfThemeFont = tfThemeFontScheme.MajorFont(msoThemeLatin) 

```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

