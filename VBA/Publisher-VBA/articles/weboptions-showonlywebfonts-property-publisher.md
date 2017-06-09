---
title: WebOptions.ShowOnlyWebFonts Property (Publisher)
keywords: vbapb10.chm8257544
f1_keywords:
- vbapb10.chm8257544
ms.prod: publisher
api_name:
- Publisher.WebOptions.ShowOnlyWebFonts
ms.assetid: d18197f4-9abe-d523-77fd-f33a8ecc8076
ms.date: 06/08/2017
---


# WebOptions.ShowOnlyWebFonts Property (Publisher)

Returns or sets a **Boolean** value that specifies whether only Web-safe fonts and font schemes should be used when the Web site is viewed in a browser. If **True**, only Web-safe fonts and font schemes are used. If  **False**, display is not limited to Web-safe fonts and font schemes. The default value is  **False**. Read/write.


## Syntax

 _expression_. **ShowOnlyWebFonts**

 _expression_A variable that represents a  **WebOptions** object.


### Return Value

Boolean


## Remarks

This property applies to Latin-based fonts only.


## Example

The following example specifies that only Web-safe fonts and font schemes should be used when the Web site is viewed in a browser.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .ShowOnlyWebFonts = True 
End With
```


