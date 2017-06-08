---
title: Document.BuiltInDocumentProperties Property (Word)
keywords: vbawd10.chm158008296
f1_keywords:
- vbawd10.chm158008296
ms.prod: word
api_name:
- Word.Document.BuiltInDocumentProperties
ms.assetid: 5e9a17dd-75b3-50e5-359e-dc0d0a59c46f
ms.date: 06/08/2017
---


# Document.BuiltInDocumentProperties Property (Word)

Returns a  **DocumentProperties** collection that represents all the built-in document properties for the specified document.


## Syntax

 _expression_ . **BuiltInDocumentProperties**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

To return a single  **DocumentProperty** object that represents a specific built-in document property, use the **BuiltinDocumentProperties** property. If Microsoft Word doesn't define a value for one of the built-in document properties, reading the **Value** property for that document property generates an error.

 For information about returning a single member of a collection, see[Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).

Use the  **CustomDocumentProperties** property to return the collection of custom document properties.


## Example

This example inserts a list of built-in properties at the end of the active document.


```vb
Sub ListProperties() 
 Dim rngDoc As Range 
 Dim proDoc As DocumentProperty 
 
 Set rngDoc = ActiveDocument.Content 
 
 rngDoc.Collapse Direction:=wdCollapseEnd 
 
 For Each proDoc In ActiveDocument.BuiltInDocumentProperties 
 With rngDoc 
 .InsertParagraphAfter 
 .InsertAfter proDoc.Name &; "= " 
 On Error Resume Next 
 .InsertAfter proDoc.Value 
 End With 
 Next 
End Sub
```

This example displays the number of words in the active document.




```vb
Sub DisplayTotalWords() 
 Dim intWords As Integer 
 intWords = ActiveDocument.BuiltInDocumentProperties(wdPropertyWords) 
 MsgBox "This document contains " &; intWords &; " words." 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

