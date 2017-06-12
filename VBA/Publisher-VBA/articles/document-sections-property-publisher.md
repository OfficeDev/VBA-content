---
title: Document.Sections Property (Publisher)
keywords: vbapb10.chm196738
f1_keywords:
- vbapb10.chm196738
ms.prod: publisher
api_name:
- Publisher.Document.Sections
ms.assetid: 9e425836-1d62-99ef-2984-b61f3a3cf831
ms.date: 06/08/2017
---


# Document.Sections Property (Publisher)

Returns a  **Sections** object representing a collection of **Section** objects in the specified document. Read-only **Sections**.


## Syntax

 _expression_. **Sections**

 _expression_A variable that represents a  **Document** object.


### Return Value

Sections


## Example

This example sets an object variable to the  **Sections** object of the active document and adds a new section starting at the second page of the publication. This example assumes that there are at least two pages in the publication.


```vb
Dim objSections As Sections 
Set objSections = ActiveDocument.Sections 
objSections.Add StartPageIndex:=2 

```


