---
title: TextFrame.HasPreviousLink Property (Publisher)
keywords: vbapb10.chm3866641
f1_keywords:
- vbapb10.chm3866641
ms.prod: publisher
api_name:
- Publisher.TextFrame.HasPreviousLink
ms.assetid: 85e0b497-55c9-d49f-2b65-e199361c121a
ms.date: 06/08/2017
---


# TextFrame.HasPreviousLink Property (Publisher)

Returns  **msoTrue** if the specified text frame has a valid link to a backward text box and **msoFalse** if it does not. Read-only.


## Syntax

 _expression_. **HasPreviousLink**

 _expression_A variable that represents a  **TextFrame** object.


### Return Value

MsoTriState


## Example

This example breaks all links in the document to the first specified text frame if links exist. This example assumes that there is at least one shape on the first page of the active publication.


```vb
Sub AddPreviousNextLinkPages() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 If .HasNextLink Then .BreakForwardLink 
 If .HasPreviousLink Then .PreviousLinkedTextFrame _ 
 .BreakForwardLink 
 End With 
End Sub
```


