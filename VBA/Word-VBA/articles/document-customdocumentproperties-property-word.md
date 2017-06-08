---
title: Document.CustomDocumentProperties Property (Word)
keywords: vbawd10.chm158007298
f1_keywords:
- vbawd10.chm158007298
ms.prod: word
api_name:
- Word.Document.CustomDocumentProperties
ms.assetid: 4f8ac449-b9b3-45a0-7962-df7252067e67
ms.date: 06/08/2017
---


# Document.CustomDocumentProperties Property (Word)

Returns a  **DocumentProperties** collection that represents all the custom document properties for the specified document.


## Syntax

 _expression_ . **CustomDocumentProperties**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use the  **BuiltInDocumentProperties** property to return the collection of built-in document properties.

Properties of type  **msoPropertyTypeString** cannot exceed 255 characters in length.

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example inserts a list of custom built-in properties at the end of the active document.


```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
For Each prop In ActiveDocument.CustomDocumentProperties 
 With myRange 
 .InsertParagraphAfter 
 .InsertAfter prop.Name &; "= " 
 .InsertAfter prop.Value 
 End With 
Next
```

This example adds a custom built-in property to Sales.doc.




```
thename = InputBox("Please type your name", "Name") 
Documents("Sales.doc").CustomDocumentProperties.Add _ 
 Name:="YourName", LinkToContent:=False, Value:=thename, _ 
 Type:=msoPropertyTypeString
```


## See also


#### Concepts


[Document Object](document-object-word.md)

