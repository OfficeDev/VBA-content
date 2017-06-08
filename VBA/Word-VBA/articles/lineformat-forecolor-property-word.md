---
title: LineFormat.ForeColor Property (Word)
keywords: vbawd10.chm164233324
f1_keywords:
- vbawd10.chm164233324
ms.prod: word
api_name:
- Word.LineFormat.ForeColor
ms.assetid: 16f8ddbe-21d8-4c09-ac54-feeb8f71146b
ms.date: 06/08/2017
---


# LineFormat.ForeColor Property (Word)

Returns or sets a  **[ColorFormat](colorformat-object-word.md)** object that represents the foreground color for the line. Read/write.


## Syntax

 _expression_ . **ForeColor**

 _expression_ A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a patterned line to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument
```


```vb
With docActive.Shapes.AddLine(10, 100, 250, 0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

