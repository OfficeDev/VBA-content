---
title: InlineShapes Object (Word)
ms.prod: word
ms.assetid: 88c632b2-80de-c96a-8879-a98461b38bd0
ms.date: 06/08/2017
---


# InlineShapes Object (Word)

A collection of  **InlineShape** objects that represent all the inline shapes in a document, range, or selection.


## Remarks

Use the  **InlineShapes** property to return the **InlineShapes** collection. The following example converts each inline shape in the active document to a **Shape** object.


```vb
For Each iShape In ActiveDocument.InlineShapes 
 iShape.ConvertToShape 
Next iShape
```

Use the  **New** method to create a new picture as an inline shape. You can use the **AddPicture** and **AddOLEObject** methods to add pictures or OLE objects and link them to a source file. Use the **AddOLEControl** method to add an ActiveX control.

 **Shape** objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. You can use the **ConvertToInlineShape** method and the **ConvertToShape** method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes.

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.

When you open a document created in an earlier version of Word, pictures are converted to inline shapes.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


