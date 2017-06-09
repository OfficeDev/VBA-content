---
title: TableOfAuthorities.Category Property (Word)
keywords: vbawd10.chm152109059
f1_keywords:
- vbawd10.chm152109059
ms.prod: word
api_name:
- Word.TableOfAuthorities.Category
ms.assetid: 29ca2198-c539-e26b-cd63-6fd5e1733e80
ms.date: 06/08/2017
---


# TableOfAuthorities.Category Property (Word)

Returns or sets the category of entries to be included in a table of authorities. Read/write  **Long** .


## Syntax

 _expression_ . **Category**

 _expression_ A variable that represents a **[TableOfAuthorities](tableofauthorities-object-word.md)** collection.


## Remarks

This property corresponds to the \c switch for a TOA field. Values 1 through 16 correspond to the items in the  **Category** list on the **Table of Authorities** tab in the **Index and Tables** dialog box. The number 0 (zero), which corresponds to all categories, cannot be used with this property. You can, however, use 0 to specify all categories when you are inserting a table of authorities.


## Example

 The following example inserts a table of authorities for all categories.


```vb
ActiveDocument.TablesOfAuthorities.Add _ 
 Range:=Selection.Range, Category:=0
```

This example formats the first table of authorities in the active document to include all citations in the first category (by default, the Cases category).




```vb
If ActiveDocument.TablesOfAuthorities.Count >= 1 Then 
 ActiveDocument.TablesOfAuthorities(1).Category = 1 
End If
```


## See also


#### Concepts


[TableOfAuthorities Object](tableofauthorities-object-word.md)

