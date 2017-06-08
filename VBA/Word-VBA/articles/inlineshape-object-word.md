---
title: InlineShape Object (Word)
keywords: vbawd10.chm2472
f1_keywords:
- vbawd10.chm2472
ms.prod: word
api_name:
- Word.InlineShape
ms.assetid: a8fd110a-4aa7-c4b9-1559-32022787d955
ms.date: 06/08/2017
---


# InlineShape Object (Word)

Represents an object in the text layer of a document. An inline shape can only be a picture, an OLE object, or an ActiveX control. The  **InlineShape** object is a member of the **[InlineShapes](inlineshapes-object-word.md)** collection. The **InlineShapes** collection contains all the shapes that appear inline in a document, range, or selection.


## Remarks

 **InlineShape** objects are treated like characters and are positioned as characters within a line of text.

Use  **InlineShapes** (Index), where Index is the index number, to return a single **InlineShape** object. Inline shapes don't have names. The following example activates the first inline shape in the active document.




```vb
ActiveDocument.InlineShapes(1).Activate
```

 **Shape** objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. You can use the **ConvertToInlineShape** method and the **ConvertToShape** method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes. Use the **Type** property to return the type of inline shape: picture, linked picture, embedded OLE object, linked OLE object, or ActiveX control.


 **Note**  When you open a document created in an earlier version of Word, pictures are converted to inline shapes.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


