---
title: Paragraph.IsStyleSeparator Property (Word)
keywords: vbawd10.chm156696710
f1_keywords:
- vbawd10.chm156696710
ms.prod: word
api_name:
- Word.Paragraph.IsStyleSeparator
ms.assetid: 7143ac54-0de8-ed70-e212-5d48b5718302
ms.date: 06/08/2017
---


# Paragraph.IsStyleSeparator Property (Word)

 **True** if a paragraph contains a special hidden paragraph mark that allows Microsoft Word to appear to join paragraphs of different paragraph styles. Read-only **Boolean** .


## Syntax

 _expression_ . **IsStyleSeparator**

 _expression_ An expression that returns a **[Paragraph](paragraph-object-word.md)** object.


## Example

This example formats all paragraphs in which there is a style separator with the built-in "Normal" style.


```vb
Sub StyleSep() 
 Dim pghDoc As Paragraph 
 For Each pghDoc In ActiveDocument.Paragraphs 
 If pghDoc.IsStyleSeparator = True Then 
 pghDoc.Range.Select 
 Selection.Style = "Normal" 
 End If 
 Next pghDoc 
End Sub
```

This example adds a paragraph after each style separator and then deletes the style separator.




```vb
Sub RemoveStyleSeparator() 
 Dim pghDoc As Paragraph 
 Dim styName As String 
 
 'Loop through all paragraphs in document to check if it is a style 
 'separator. If it is, delete it and enter a regular paragraph 
 For Each pghDoc In ActiveDocument.Paragraphs 
 If pghDoc.IsStyleSeparator = True Then 
 pghDoc.Range.Select 
 With Selection 
 .Collapse (wdCollapseEnd) 
 .TypeParagraph 
 .MoveLeft (1) 
 .TypeBackspace 
 End With 
 End If 
 Next pghDoc 
End Sub
```


## See also


#### Concepts


[Paragraph Object](paragraph-object-word.md)

