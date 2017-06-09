---
title: Font.NumberForm Property (Word)
keywords: vbawd10.chm156369071
f1_keywords:
- vbawd10.chm156369071
ms.prod: word
api_name:
- Word.Font.NumberForm
ms.assetid: 730ce7a1-a0f4-3ed8-d7a0-5b4039f56817
ms.date: 06/08/2017
---


# Font.NumberForm Property (Word)

Returns or sets the number form setting for an OpenType font. Read/write [WdNumberForm](wdnumberform-enumeration-word.md).


## Syntax

 _expression_ . **NumberForm**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Numbers in OpenType fonts can be displayed either with consistent heights along the baseline of the text (called "lining"), or with varying heights (called "hanging" or "old style") where numbers are displayed above or below the baseline of the text. Use the  **NumberForm** property to specify whether numbers are displayed using lining or old-style.

Setting this property has the same effect as selecting an item in the dropdown box next to  **Number forms:** ( **OpenType features** group, **Advanced** tab in the **Font** dialog in Word).


## Example

The following code example sets the number form to "Old-style" for the font in the active document.


```vb
ActiveDocument.Range.Font.NumberForm = wdNumberFormOldStyle
```


## See also


#### Concepts


[Font Object](font-object-word.md)

