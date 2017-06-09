---
title: ChartFont.Bold Property (PowerPoint)
keywords: vbapp10.chm704002
f1_keywords:
- vbapp10.chm704002
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.Bold
ms.assetid: 5d5a0b2e-5aab-f197-79da-e9bb8d219af9
ms.date: 06/08/2017
---


# ChartFont.Bold Property (PowerPoint)

 **True** if the font is bold. Read/write **Variant**.


## Syntax

 _expression_. **Bold**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font to bold for all characters in the chart title of the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.ChartTitle.Characters.Font.Bold = True

    End If

End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-powerpoint.md)

