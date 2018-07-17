---
title: Paragraph.SelectNumber Method (Word)
keywords: vbawd10.chm156696911
f1_keywords:
- vbawd10.chm156696911
ms.prod: word
api_name:
- Word.Paragraph.SelectNumber
ms.assetid: 9b5999d4-da07-8a32-4aa9-9b62f9cd9e31
ms.date: 06/08/2017
---


# Paragraph.SelectNumber Method (Word)

Selects the number or bullet in a list.


## Syntax

 _expression_ . **SelectNumber**

 _expression_ Required. A variable that represents a **[Paragraph](paragraph-object-word.md)** object.


## Remarks

If the  **SelectNumber** method is called from a paragraph, selection, or range that does not contain a list, an error message is displayed.


## Example

This example selects the bullet or number for the list that contains the selected paragraph in the active document, and then it increases the font size of the bullet or number to 17 points. This example assumes that the first paragraph in the selection is formatted as a bulleted or numbered list.


```vb
Sub SelectParaNumber() 
 With Selection 
 .Paragraphs(1).SelectNumber 
 .Font.Size = 17 
 End With 
End Sub
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

