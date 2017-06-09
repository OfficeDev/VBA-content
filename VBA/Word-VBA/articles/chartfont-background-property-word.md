---
title: ChartFont.Background Property (Word)
keywords: vbawd10.chm255918080
f1_keywords:
- vbawd10.chm255918080
ms.prod: word
api_name:
- Word.ChartFont.Background
ms.assetid: 3ae75226-265d-f544-489d-e3e417995ef8
ms.date: 06/08/2017
---


# ChartFont.Background Property (Word)

Returns or sets the type of background for text used in charts. Read/write  **Variant** that is set to one of the constants of **[XlBackground](xlbackground-enumeration-word.md)** .


## Syntax

 _expression_ . **Background**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-word.md)** object.


## Example

The following example adds a chart title to the first chart in the active document and then sets the font size and specifies a transparent background for the title.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasTitle = True 
 .ChartTitle.Text = "Rainfall Totals by Month" 
 With .ChartTitle.Font 
 .Size = 10 
 .Background = xlBackgroundTransparent 
 End With 
 End With 
 End If 
End With
```


## See also


#### Concepts


[ChartFont Object](chartfont-object-word.md)

