---
title: Font.ItalicBi Property (Word)
keywords: vbawd10.chm156369057
f1_keywords:
- vbawd10.chm156369057
ms.prod: word
api_name:
- Word.Font.ItalicBi
ms.assetid: 56b1a7cb-2e42-7ff7-d7b8-80f047fb3d4b
ms.date: 06/08/2017
---


# Font.ItalicBi Property (Word)

 **True** if the font or range is formatted as italic. Read/write **Long** .


## Syntax

 _expression_ . **ItalicBi**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

This property returns  **True** , **False** or **wdUndefined** (for a mixture of italic and non-italic text) and can be set to **True** , **False** , or **wdToggle** .

Use the  **ItalicBi** property for right-to-left languages.


## Example

This example italicizes the first paragraph in the active right-to-left language document.


```vb
ActiveDocument.Paragraphs(1).Range.ItalicBi = True
```


## See also


#### Concepts


[Font Object](font-object-word.md)

