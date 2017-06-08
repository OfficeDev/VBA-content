---
title: Document.TablesOfAuthoritiesCategories Property (Word)
keywords: vbawd10.chm158007334
f1_keywords:
- vbawd10.chm158007334
ms.prod: word
api_name:
- Word.Document.TablesOfAuthoritiesCategories
ms.assetid: c7daaf7a-6002-8377-ff68-18335f441baf
ms.date: 06/08/2017
---


# Document.TablesOfAuthoritiesCategories Property (Word)

Returns a  **[TablesOfAuthoritiesCategories](tablesofauthoritiescategories-object-word.md)** collection that represents the available table of authorities categories for the specified document. Read-only.


## Syntax

 _expression_ . **TablesOfAuthoritiesCategories**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example changes the name of the eighth item in the table of authorities category list for the active document.


```vb
ActiveDocument.TablesOfAuthoritiesCategories(8).Name = "Other case"
```

This example displays the name of the last table of authorities category if the category name has been changed.




```vb
last = ActiveDocument.TablesOfAuthoritiesCategories.Count 
If ActiveDocument.TablesOfAuthoritiesCategories(last) _ 
 .Name <> "16" Then 
 MsgBox ActiveDocument.TablesOfAuthoritiesCategories(last).Name 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

