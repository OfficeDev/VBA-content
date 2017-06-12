---
title: Global.Documents Property (Word)
keywords: vbawd10.chm163119105
f1_keywords:
- vbawd10.chm163119105
ms.prod: word
api_name:
- Word.Global.Documents
ms.assetid: a86bad22-aabf-dd0d-4b23-fc608d5db4c1
ms.date: 06/08/2017
---


# Global.Documents Property (Word)

Returns a  **[Documents](documents-object-word.md)** collection that represents all the open documents. Read-only.


## Syntax

 _expression_ . **Documents**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example creates a new document based on the Normal template and then displays the Save As dialog box.


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


[Global Object](global-object-word.md)

