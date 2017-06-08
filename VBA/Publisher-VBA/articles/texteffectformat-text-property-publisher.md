---
title: TextEffectFormat.Text Property (Publisher)
keywords: vbapb10.chm3735824
f1_keywords:
- vbapb10.chm3735824
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat.Text
ms.assetid: eae1e95f-b0e6-559b-39a5-40291e758915
ms.date: 06/08/2017
---


# TextEffectFormat.Text Property (Publisher)

Returns or sets a  **String** that represents the text in a text range or WordArt shape. Read/write.


## Syntax

 _expression_. **Text**

 _expression_A variable that represents a  **TextEffectFormat** object.


## Example

The following example changes the text and sets the font name and formatting properties for shape one on the first page of the active publication. For this example to work, shape one must be a WordArt object.


```vb
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .Text = "This is a test." 
 .FontName = "Courier New" 
 .FontBold = True 
 .FontItalic = True 
 End With 
End Sub
```


