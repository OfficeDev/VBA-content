---
title: ChartColorFormat.SchemeColor Property (Word)
keywords: vbawd10.chm12060270
f1_keywords:
- vbawd10.chm12060270
ms.prod: word
api_name:
- Word.ChartColorFormat.SchemeColor
ms.assetid: 56832016-dcd9-5627-d0e4-8cce040c24f7
ms.date: 06/08/2017
---


# ChartColorFormat.SchemeColor Property (Word)

Returns or sets the index of a color in the current color scheme. Read/write  **Long** .


## Syntax

 _expression_ . **SchemeColor**

 _expression_ A variable that represents a **[ChartColorFormat](chartcolorformat-object-word.md)** object.


## Example

The following example sets the visibility, foreground color, background color, and gradient for the chart area fill of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
 End With 
 End If 
End With 

```


## See also


#### Concepts


[ChartColorFormat Object](chartcolorformat-object-word.md)

