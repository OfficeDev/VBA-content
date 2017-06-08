---
title: Chart.DisplayBlanksAs Property (PowerPoint)
keywords: vbapp10.chm684026
f1_keywords:
- vbapp10.chm684026
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.DisplayBlanksAs
ms.assetid: 8f00f6dc-3885-1f97-057d-3c426c19a1a1
ms.date: 06/08/2017
---


# Chart.DisplayBlanksAs Property (PowerPoint)

Returns or sets the way that blank cells are plotted on a chart. Can be one of the  **[XlDisplayBlanksAs](xldisplayblanksas-enumeration-powerpoint.md)** constants. Read/write **Long**.


## Syntax

 _expression_. **DisplayBlanksAs**

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets Microsoft Word to not plot blank cells for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.DisplayBlanksAs = xlNotPlotted

    End If

End With
```


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

