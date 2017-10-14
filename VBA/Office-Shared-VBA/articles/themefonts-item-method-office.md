---
title: ThemeFonts.Item Method (Office)
ms.prod: office
api_name:
- Office.ThemeFonts.Item
ms.assetid: 09b437dd-9be3-223e-4b81-f83a1d44d53f
ms.date: 06/08/2017
---


# ThemeFonts.Item Method (Office)

Gets one of the three language fonts contained in the  **ThemeFonts** collection.


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ An expression that returns a **ThemeFonts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**MsoFontLanguageIndex**|The index value of the  **ThemeFont** object.|

### Return Value

ThemeFont


## Example

The following example sets the font for the body of a document to the Latin theme.


```
Dim tTheme As OfficeTheme 
Dim tfThemeFonts As ThemeFonts 
Dim latinMinorFont As ThemeFont 
Set tfThemeFonts = tTheme.ThemeFontScheme.MinorFont 
Set latinMinorFont = tfThemeFonts(msoThemeLatin)
```


## See also


#### Concepts


[ThemeFonts Object](themefonts-object-office.md)
#### Other resources


[ThemeFonts Object Members](themefonts-members-office.md)

