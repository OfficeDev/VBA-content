---
title: Font.BoldBi Property (Word)
keywords: vbawd10.chm156369056
f1_keywords:
- vbawd10.chm156369056
ms.prod: word
api_name:
- Word.Font.BoldBi
ms.assetid: 75c49bb4-acc7-17d7-5887-f7ecf87dd5df
ms.date: 06/08/2017
---


# Font.BoldBi Property (Word)

 **True** if the font is formatted as bold. Read/write **Long** .


## Syntax

 _expression_ . **BoldBi**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

This property returns  **True** , **False** or **wdUndefined** (for a mixture of bold and non-bold text). Can be set to **True** , **False** , or **wdToggle** .

The  **BoldBi** property applies to text in a right-to-left language.


## Example

This example makes the first paragraph in the active right-to-left language document bold.


```vb
ActiveDocument.Paragraphs(1).Range.Font.BoldBi = True
```


## See also


#### Concepts


[Font Object](font-object-word.md)

