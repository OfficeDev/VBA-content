---
title: WebPageFont.FixedWidthFontSize Property (Office)
keywords: vbaof11.chm224004
f1_keywords:
- vbaof11.chm224004
ms.prod: office
api_name:
- Office.WebPageFont.FixedWidthFontSize
ms.assetid: a3f68d85-219d-c94b-15d2-c55374158fc2
ms.date: 06/08/2017
---


# WebPageFont.FixedWidthFontSize Property (Office)

Sets or gets the fixed-width font size setting in the host application, in points. Read/write.


## Syntax

 _expression_. **FixedWidthFontSize**

 _expression_ A variable that represents a **WebPageFont** object.


## Remarks

When you set the  **FixedWidthFontSize** property, the host application does not check the value for validity. If you enter an invalid value, such as a nonnumber, the host application sets the size to 0 points. You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


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

