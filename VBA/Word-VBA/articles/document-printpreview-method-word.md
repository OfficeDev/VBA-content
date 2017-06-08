---
title: Document.PrintPreview Method (Word)
keywords: vbawd10.chm158007410
f1_keywords:
- vbawd10.chm158007410
ms.prod: word
api_name:
- Word.Document.PrintPreview
ms.assetid: 534e3a03-b26c-5144-f6f5-09235830ec4f
ms.date: 06/08/2017
---


# Document.PrintPreview Method (Word)

Switches the view to print preview.


## Syntax

 _expression_ . **PrintPreview**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

In addition to using the  **PrintPreview** method, you can set the **[PrintPreview](application-printpreview-property-word.md)** property to **True** or **False** to switch to or from print preview, respectively. You can also change the view by setting the **[Type](document-type-property-word.md)** property for the **View** object to **wdPrintPreview** .


## Example

This example switches the active document to print preview if it is currently in some other view.


```vb
If Application.PrintPreview = False Then 
 ActiveDocument.PrintPreview 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

