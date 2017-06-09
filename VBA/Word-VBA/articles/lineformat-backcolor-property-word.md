---
title: LineFormat.BackColor Property (Word)
keywords: vbawd10.chm164233316
f1_keywords:
- vbawd10.chm164233316
ms.prod: word
api_name:
- Word.LineFormat.BackColor
ms.assetid: 30c282ca-f20b-9d20-8e6c-6f2c37d0cd7b
ms.date: 06/08/2017
---


# LineFormat.BackColor Property (Word)

Returns or sets a  **[ColorFormat](colorformat-object-word.md)** object that represents the background color for a patterned line. Read/write.


## Syntax

 _expression_ . **BackColor**

 _expression_ A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a patterned line to the active document.


```vb
Dim docActive As Document 
 
Set docActive = ActiveDocument 
 
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

