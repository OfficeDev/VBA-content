---
title: LineFormat.Pattern Property (Word)
keywords: vbawd10.chm164233325
f1_keywords:
- vbawd10.chm164233325
ms.prod: word
api_name:
- Word.LineFormat.Pattern
ms.assetid: 6aa5b1e1-813c-bf03-aafa-7ef2aacbe51e
ms.date: 06/08/2017
---


# LineFormat.Pattern Property (Word)

Returns or sets a value that represents the pattern applied to the specified line. Read/write  **MsoPatternType** .


## Syntax

 _expression_ . **Pattern**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a patterned line to  _myDocument_ .


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddLine(10, 100, 250, 0).Line 
 .Weight = 6 
 .ForeColor.RGB = RGB(0, 0, 255) 
 .BackColor.RGB = RGB(128, 0, 0) 
 .Pattern = msoPatternDarkDownwardDiagonal 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

