---
title: Shape.ConvertToInlineShape Method (Word)
keywords: vbawd10.chm161480729
f1_keywords:
- vbawd10.chm161480729
ms.prod: word
api_name:
- Word.Shape.ConvertToInlineShape
ms.assetid: 367b6d36-dd62-aa2b-62df-d5f42ac2699c
ms.date: 06/08/2017
---


# Shape.ConvertToInlineShape Method (Word)

Converts the specified shape in the drawing layer of a document to an inline shape in the text layer. You can convert only shapes that represent pictures, OLE objects, or ActiveX controls. This method returns an  **[InlineShape](inlineshape-object-word.md)** object that represents the picture or OLE object.


## Syntax

 _expression_ . **ConvertToInlineShape**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

Shapes that support attached text cannot be converted to inline shapes. For these shapes, use the  **ConvertToFrame** method.



If you use this method on a  **ShapeRange** object that contains more than one shape, an error occurs.




## Example

This example converts each picture in MyDoc.doc to an inline shape.


```vb
For Each s In Documents("MyDoc.doc").Shapes 
 If s.Type = msoPicture Then 
 s.ConvertToInlineShape 
 End If 
Next s
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

