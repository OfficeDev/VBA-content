---
title: ChartFont.FontStyle Property (PowerPoint)
keywords: vbapp10.chm704005
f1_keywords:
- vbapp10.chm704005
ms.prod: powerpoint
api_name:
- PowerPoint.ChartFont.FontStyle
ms.assetid: b93a278e-cf38-ef2a-acdc-862fc4ca0b1c
ms.date: 06/08/2017
---


# ChartFont.FontStyle Property (PowerPoint)

Returns or sets the font style. Read/write  **String**.


## Syntax

 _expression_. **FontStyle**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-powerpoint.md)** object.


## Remarks

Changing this property may affect other  **ChartFont** properties (such as **[Bold](chartfont-bold-property-powerpoint.md)** and **[Italic](chartfont-italic-property-powerpoint.md)** ).


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the font style for the title of the first chart in the active document to bold and italic.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Title.Font.FontStyle = "Bold Italic"

    End If

End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-powerpoint.md)

