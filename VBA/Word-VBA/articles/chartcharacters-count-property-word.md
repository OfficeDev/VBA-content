---
title: ChartCharacters.Count Property (Word)
keywords: vbawd10.chm250740854
f1_keywords:
- vbawd10.chm250740854
ms.prod: word
api_name:
- Word.ChartCharacters.Count
ms.assetid: 8ee2abf3-4d80-a235-8fbc-a011842da718
ms.date: 06/08/2017
---


# ChartCharacters.Count Property (Word)

Returns the number of objects in the collection. Read-only  **Long** .


## Syntax

 _expression_ . **Count**

 _expression_ A variable that represents a **[ChartCharacters](chartcharacters-object-word.md)** object.


## Example

The following example makes the last character a superscript character in the title of the first chart in the active document.


```vb
Sub MakeSuperscript() 
 Dim n As Integer 
 
 With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 n = .Chart.Title.Characters.Count 
 .Chart.Title.Characters(n, 1).Font.Superscript = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[ChartCharacters Object](chartcharacters-object-word.md)

