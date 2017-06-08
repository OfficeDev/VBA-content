---
title: ThemeFonts Object (Office)
ms.prod: office
api_name:
- Office.ThemeFonts
ms.assetid: 393865af-f008-d26c-5b82-9ae79766e511
ms.date: 06/08/2017
---


# ThemeFonts Object (Office)

Represents a collection of major and minor fonts in the font scheme of a Microsoft Office theme.


## Example

The following example sets a  **ThemeFonts** object to a minor theme font.


```
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 

```


## Methods



|**Name**|
|:-----|
|[Item](themefonts-item-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](themefonts-application-property-office.md)|
|[Count](themefonts-count-property-office.md)|
|[Creator](themefonts-creator-property-office.md)|
|[Parent](themefonts-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
