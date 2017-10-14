---
title: ThemeFont Object (Office)
ms.prod: office
api_name:
- Office.ThemeFont
ms.assetid: 1a9f1365-c392-3d04-74db-333ac111114a
ms.date: 06/08/2017
---


# ThemeFont Object (Office)

Represents a container for the font schemes of a Microsoft Office theme.


## Example

The following example sets the Headings font scheme in a Microsoft Office theme to a Latin scheme. 


```
Dim tTheme As OfficeTheme 
Dim tfThemeFontScheme As ThemeFontScheme 
Dim tfThemeFont As ThemeFont 
Set tfThemeFontScheme = tTheme.ThemeFontScheme 
Set tfThemeFont = tfThemeFontScheme.MajorFont(msoThemeLatin) 

```


## Properties



|**Name**|
|:-----|
|[Application](themefont-application-property-office.md)|
|[Creator](themefont-creator-property-office.md)|
|[Name](themefont-name-property-office.md)|
|[Parent](themefont-parent-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
