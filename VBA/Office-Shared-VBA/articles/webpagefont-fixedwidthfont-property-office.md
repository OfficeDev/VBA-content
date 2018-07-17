---
title: WebPageFont.FixedWidthFont Property (Office)
keywords: vbaof11.chm224003
f1_keywords:
- vbaof11.chm224003
ms.prod: office
api_name:
- Office.WebPageFont.FixedWidthFont
ms.assetid: f522922a-097f-2b94-42cf-680393e513b9
ms.date: 06/08/2017
---


# WebPageFont.FixedWidthFont Property (Office)

Sets or gets the fixed-width font setting in the host application. Read/write.


## Syntax

 _expression_. **FixedWidthFont**

 _expression_ A variable that represents a **WebPageFont** object.


## Remarks

When you set the  **FixedWidthFont** property, the host application does not check the value for validity.


## Example

This example sets the fixed-width font and fixed-width font size for the English/Western European/Other Latin Script character set in the active application.


```
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.FixedWidthFont = "System" 
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.FixedWidthFontSize = 12
```


## See also


#### Concepts


[WebPageFont Object](webpagefont-object-office.md)
#### Other resources


[WebPageFont Object Members](webpagefont-members-office.md)

