---
title: TextFrame.HasNextLink Property (Publisher)
keywords: vbapb10.chm3866640
f1_keywords:
- vbapb10.chm3866640
ms.prod: publisher
api_name:
- Publisher.TextFrame.HasNextLink
ms.assetid: 907ec470-e283-906a-e25f-f5a8548a18a4
ms.date: 06/08/2017
---


# TextFrame.HasNextLink Property (Publisher)

Indicates whether the specified text frame has a valid forward text-box link. Read-only.


## Syntax

 _expression_. **HasNextLink**

 _expression_A variable that represents a  **TextFrame** object.


### Return Value

MsoTriState


## Remarks

The  **HasNextLink** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The specified text frame does not have a forward text-box link.|
| **msoTrue**| The specified text frame has a forward text-box link.|

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


