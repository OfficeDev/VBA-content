---
title: ShapeRange.Hyperlink Property (Word)
keywords: vbawd10.chm162857961
f1_keywords:
- vbawd10.chm162857961
ms.prod: word
api_name:
- Word.ShapeRange.Hyperlink
ms.assetid: a9b5176d-932c-b7b9-be56-ece4240bbf35
ms.date: 06/08/2017
---


# ShapeRange.Hyperlink Property (Word)

Returns a  **Hyperlink** object that represents the hyperlink associated with the specified **ShapeRange** object. Read-only.


## Syntax

 _expression_ . **Hyperlink**

 _expression_ A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

If there is no hyperlink associated with the specified range of shapes, an error occurs. In this case, use the  **[Add](hyperlinks-add-method-word.md)** method for the **[Hyperlinks](hyperlinks-object-word.md)** collection to add a hyperlink to the specified range of shapes. The following example shows how to do this.


```vb
ActiveDocument.Hyperlinks.Add Selection.ShapeRange(1), "http://www.microsoft.com"
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

