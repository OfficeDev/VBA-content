---
title: Documents Object (Word)
ms.prod: word
ms.assetid: fc4ac973-19c1-703a-5538-f4426b8b7564
ms.date: 06/08/2017
---


# Documents Object (Word)

A collection of all the  **[Document](document-object-word.md)** objects that are currently open in Word.


## Remarks

Use the  **Documents** property to return the **Documents** collection. The following example displays the names of the open documents.


```vb
For Each aDoc In Documents 
 aName = aName &; aDoc.Name &; vbCr 
Next aDoc 
MsgBox aName
```

Use the  **[Add](documents-add-method-word.md)** method to create a new empty document and add it to the **Documents** collection. The following example creates a new document based on the Normal template.




```
Documents.Add
```

Use the  **[Open](documents-open-method-word.md)** method to open a file. The following example opens the document named "Sales.doc."




```
Documents.Open FileName:="C:\My Documents\Sales.doc"
```

Use  **[Documents](application-documents-property-word.md)** (Index), where Index is the document name or index number to return a single **Document** object. The following instruction closes the document named "Report.doc" without saving changes.




```
Documents("Report.doc").Close SaveChanges:=wdDoNotSaveChanges
```

The index number represents the position of the document in the  **Documents** collection. The following example activates the first document in the **Documents** collection.




```
Documents(1).Activate
```

The following example enumerates the  **Documents** collection to determine whether the document named "Report.doc" is open. If this document is contained in the **Documents** collection, the document is activated; otherwise, it is opened.




```vb
For Each doc In Documents 
 If doc.Name = "Report.doc" Then found = True 
Next doc 
If found <> True Then 
 Documents.Open FileName:="C:\Documents\Report.doc" 
Else 
 Documents("Report.doc").Activate 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

