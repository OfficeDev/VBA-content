---
title: TextFrame.DeleteText Method (PowerPoint)
keywords: vbapp10.chm558014
f1_keywords:
- vbapp10.chm558014
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.DeleteText
ms.assetid: 0971765b-8d2c-a34a-7184-119af42be835
ms.date: 06/08/2017
---


# TextFrame.DeleteText Method (PowerPoint)

Deletes the text associated with the specified shape.


## Syntax

 _expression_. **DeleteText**

 _expression_ A variable that represents a **TextFrame** object.


## Example

If shape two on  `myDocument` contains text, this example deletes the text.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(2).TextFrame.DeleteText
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

