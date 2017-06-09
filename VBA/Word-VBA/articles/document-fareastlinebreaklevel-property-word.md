---
title: Document.FarEastLineBreakLevel Property (Word)
keywords: vbawd10.chm158007607
f1_keywords:
- vbawd10.chm158007607
ms.prod: word
api_name:
- Word.Document.FarEastLineBreakLevel
ms.assetid: 11642adb-2c15-a081-ae7c-d9ebe6d5b848
ms.date: 06/08/2017
---


# Document.FarEastLineBreakLevel Property (Word)

Returns or sets a  **WdFarEastLineBreakLevel** that represents the line break control level for the specified document. Read/write.


## Syntax

 _expression_ . **FarEastLineBreakLevel**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

This property is ignored if the  **FarEastLineBreakControl** property is set to **False** .

For more information on using Microsoft Word with East Asian languages, see Word features for East Asian languages .


## Example

This example sets Microsoft Word to perform line breaking on first-level kinsoku characters in the active document.


```vb
ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana
```


## See also


#### Concepts


[Document Object](document-object-word.md)

