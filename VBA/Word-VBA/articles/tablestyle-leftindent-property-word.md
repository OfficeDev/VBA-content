---
title: TableStyle.LeftIndent Property (Word)
keywords: vbawd10.chm244776974
f1_keywords:
- vbawd10.chm244776974
ms.prod: word
api_name:
- Word.TableStyle.LeftIndent
ms.assetid: 5dc6a39f-ed73-8492-7ef5-b02f0290ddbc
ms.date: 06/08/2017
---


# TableStyle.LeftIndent Property (Word)

Returns or sets a  **Single** that represents the left indent value (in points) for the rows in the specified table style. Read/write.


## Syntax

 _expression_ . **LeftIndent**

 _expression_ A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


## Example

This example sets the left indent of the first paragraph in the active document to 1 inch. The  **InchesToPoints** method is used to convert inches to points.


```vb
ActiveDocument.Paragraphs(1).LeftIndent = InchesToPoints(1)
```

This example sets the left indent for all rows in the first table in the active document.




```vb
ActiveDocument.Tables(1).Rows.LeftIndent = InchesToPoints(1)
```


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

