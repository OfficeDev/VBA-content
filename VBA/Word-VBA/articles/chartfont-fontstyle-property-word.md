---
title: ChartFont.FontStyle Property (Word)
keywords: vbawd10.chm255918088
f1_keywords:
- vbawd10.chm255918088
ms.prod: word
api_name:
- Word.ChartFont.FontStyle
ms.assetid: dc877dd3-6f58-08f9-719c-2fb9edbb868f
ms.date: 06/08/2017
---


# ChartFont.FontStyle Property (Word)

Returns or sets the font style. Read/write  **String** .


## Syntax

 _expression_ . **FontStyle**

 _expression_ A variable that represents a **[ChartFont](chartfont-object-word.md)** object.


## Remarks

Changing this property may affect other  **ChartFont** properties (such as **[Bold](chartfont-bold-property-word.md)** and **[Italic](chartfont-italic-property-word.md)** ).


## Example

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


[ChartFont Object](chartfont-object-word.md)

