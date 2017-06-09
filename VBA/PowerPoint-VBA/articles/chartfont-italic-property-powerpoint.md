---
title: ChartFont.Italic Property (PowerPoint)
keywords: vbapp10.chm704006
f1_keywords:
- vbapp10.chm704006
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Italic
ms.assetid: c62ad4c5-c7b3-58d8-8d37-540a8a123ce2
ms.date: 06/08/2017
---


# ChartFont.Italic Property (PowerPoint)

 **True** if the font style is italic. Read/write **Boolean**.


## Syntax

 _expression_. **Italic**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[ChartFont Object](chartfont-object-powerpoint.md)

