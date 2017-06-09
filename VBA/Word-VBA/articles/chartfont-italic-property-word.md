---
title: ChartFont.Italic Property (Word)
keywords: vbawd10.chm255918090
f1_keywords:
- vbawd10.chm255918090
ms.prod: word
api_name:
- Word.ChartFont.Italic
ms.assetid: 8e25a2dd-2ac1-83ec-c505-fdc23b0de7d9
ms.date: 06/08/2017
---


# ChartFont.Italic Property (Word)

 **True** if the font style is italic. Read/write **Boolean** .


## Syntax

 _expression_ . **Italic**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-word.md)** object.


## Example

The following example sets the font to italic for all characters in the title of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Title.Characters.Font.Italic = True 
 End If 
End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-word.md)

