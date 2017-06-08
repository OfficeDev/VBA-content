---
title: Application.Documents Property (Word)
keywords: vbawd10.chm158334982
f1_keywords:
- vbawd10.chm158334982
ms.prod: word
api_name:
- Word.Application.Documents
ms.assetid: 7e477cb3-ae65-685a-0083-1826efe86703
ms.date: 06/08/2017
---


# Application.Documents Property (Word)

Returns a  **[Documents](documents-object-word.md)** collection that represents all the open documents. Read-only.


## Syntax

 _expression_ . **Documents**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


 **Note**  A document displayed in a protected view window is not a member of the  **[Documents](application-documents-property-word.md)** collection. Instead, use the[Document](document-object-word.md) property of the[ProtectedViewWindow](protectedviewwindow-object-word.md) object to access a document that is displayed in a protected view window.


## Example

This example creates a new document based on the Normal template and then displays the  **Save As** dialog box.


```
Documents.Add.Save
```

This example saves open documents that have changed since they were last saved.




```vb
Dim docLoop As Document 
 
For Each docLoop In Documents 
   If docLoop.Saved = False Then docLoop.Save 
Next docLoop
```

This example prints each open document after setting the left and right margins to 0.5 inch.




```vb
Dim docLoop As Document 
 
For Each docLoop In Documents 
    With docLoop 
        .PageSetup.LeftMargin = InchesToPoints(0.5) 
        .PageSetup.RightMargin = InchesToPoints(0.5) 
        .PrintOut 
    End With 
Next docLoop
```

This example opens Doc.doc as a read-only document.




```
Documents.Open FileName:="C:\Files\Doc.doc", ReadOnly:=True
```


## See also


#### Concepts


[Application Object](application-object-word.md)

