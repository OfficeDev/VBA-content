---
title: ShapeRange.AddToCatalogMergeArea Method (Publisher)
keywords: vbapb10.chm2294048
f1_keywords:
- vbapb10.chm2294048
ms.prod: publisher
api_name:
- Publisher.ShapeRange.AddToCatalogMergeArea
ms.assetid: 6cb770c6-fe6e-ffe8-cd51-855d97b17aed
ms.date: 06/08/2017
---


# ShapeRange.AddToCatalogMergeArea Method (Publisher)

Adds the specified shape or shapes to the publication page's catalog merge area.


## Syntax

 _expression_. **AddToCatalogMergeArea**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The catalog merge area is automatically resized to accommodate objects that are larger than the merge area, or that are positioned outside the catalog merge area when they are added.

The  **AddToCatalogMergeArea** method does not apply to merge data fields:


- Use the  **[Insert](mailmergedatafield-insert-method-publisher.md)** method of the **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** collection to add a picture data field to a publication page's catalog merge area.
    
- Use the  **[InsertMailMergeField](textrange-insertmailmergefield-method-publisher.md)** method of the **[TextRange](textrange-object-publisher.md)** object to add a text data field to a text box.
    


Note that to add a text box that will contain text data fields to a catalog merge area, you use the  **AddToCatalogMergeArea** method.


## Example

The following example adds a rectangle to the catalog merge area on the first page of the specified publication. This example assumes a catalog merge area has been added to the first page.


```vb
ThisDocument.Pages(1).Shapes.AddShape(1, 80, 75, 450, 125).AddToCatalogMergeArea
```


