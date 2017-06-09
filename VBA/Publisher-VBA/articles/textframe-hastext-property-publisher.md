---
title: TextFrame.HasText Property (Publisher)
keywords: vbapb10.chm3866642
f1_keywords:
- vbapb10.chm3866642
ms.prod: publisher
api_name:
- Publisher.TextFrame.HasText
ms.assetid: f8d1c660-c3f1-e835-adc3-114e6611de98
ms.date: 06/08/2017
---


# TextFrame.HasText Property (Publisher)

Returns an  **MsoTriState** constant indicating whether the specified shape has text associated with it. Read-only.


## Syntax

 _expression_. **HasText**

 _expression_A variable that represents a  **TextFrame** object.


## Example

If shape two on the first page of the active publication contains text, this example resizes the shape to fit the text.


```vb
With ActiveDocument.Pages(1).Shapes(2).TextFrame 
 If .HasText Then .AutoFitText = pbTextAutoFitBestFit 
End With
```


