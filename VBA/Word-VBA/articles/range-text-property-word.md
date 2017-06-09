---
title: Range.Text Property (Word)
keywords: vbawd10.chm157155328
f1_keywords:
- vbawd10.chm157155328
ms.prod: word
api_name:
- Word.Range.Text
ms.assetid: 495fe06e-ba87-0d96-9f6e-3e62fd71d4a5
ms.date: 06/08/2017
---


# Range.Text Property (Word)

Returns or sets the text in the specified range or selection. Read/write  **String** . Read/write **String** .


## Syntax

 _expression_ . **Text**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The  **Text** property returns the plain, unformatted text of the range. When you set this property, the existing text in the range is replaced.


## Example

This example replaces the first word in the active document with "Dear."


```vb
Set myRange = ActiveDocument.Words(1) 
myRange.Text = "Dear "
```


## See also


#### Concepts


[Range Object](range-object-word.md)

