---
title: Paragraph.KeepWithNext Property (Word)
keywords: vbawd10.chm156696679
f1_keywords:
- vbawd10.chm156696679
ms.prod: word
api_name:
- Word.Paragraph.KeepWithNext
ms.assetid: 59991695-23cc-9580-5a49-3e2c266938f3
ms.date: 06/08/2017
---


# Paragraph.KeepWithNext Property (Word)

 **True** if the specified paragraph remains on the same page as the paragraph that follows it when Microsoft Word repaginates the document. Read/write **Long** .


## Syntax

 _expression_ . **KeepWithNext**

 _expression_ A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** .


## Example

This example keeps the third paragraph through sixth paragraph in the active document on the same page.


```vb
For i = 3 To 5 
 ActiveDocument.Paragraphs(i).KeepWithNext = True 
Next i
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

