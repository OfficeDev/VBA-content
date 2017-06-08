---
title: WebPageFont.ProportionalFontSize Property (Office)
keywords: vbaof11.chm224002
f1_keywords:
- vbaof11.chm224002
ms.prod: office
api_name:
- Office.WebPageFont.ProportionalFontSize
ms.assetid: b51333ff-5017-8533-ea74-3a104ed67dd8
ms.date: 06/08/2017
---


# WebPageFont.ProportionalFontSize Property (Office)

Sets or gets the proportional font size setting in the host application, in points. Read/write.


## Syntax

 _expression_. **ProportionalFontSize**

 _expression_ A variable that represents a **WebPageFont** object.


## Remarks

When you set the  **ProportionalFontSize** property, the host application does not check the value for validity. If you enter an invalid value, such as a nonnumber, the host application sets the size to 0 points. You can enter half-point sizes; if you enter other fractional point sizes, they are rounded up or down to the nearest half-point.


## Example

This example sets the proportional font and proportional font size for the English/Western European/Other Latin Script character set in the active application.


```
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFont = "Tahoma" 
Application.DefaultWebOptions. _ 
Fonts(msoCharacterSetEnglishWesternEuropeanOtherLatinScript) _ 
.ProportionalFontSize = 14.5
```


## See also


#### Concepts


[WebPageFont Object](webpagefont-object-office.md)
#### Other resources


[WebPageFont Object Members](webpagefont-members-office.md)

