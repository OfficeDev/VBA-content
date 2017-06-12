---
title: Document.ScratchArea Property (Publisher)
keywords: vbapb10.chm196657
f1_keywords:
- vbapb10.chm196657
ms.prod: publisher
api_name:
- Publisher.Document.ScratchArea
ms.assetid: 782d9b7f-b620-60f0-c21d-04f588c37cc6
ms.date: 06/08/2017
---


# Document.ScratchArea Property (Publisher)

Returns a  **[ScratchArea](scratcharea-object-publisher.md)** object for an a given document.


## Syntax

 _expression_. **ScratchArea**

 _expression_A variable that represents a  **Document** object.


### Return Value

ScratchArea


## Remarks

The  **ScratchArea** object is a collection of objects on the scratch page. The **ScratchArea** object is not in the **Pages** collection because it is fundamentally not a page; its only similarity to a page is that it can contain objects.


## Example

This example sets the variable object as the first shape on the scratch area of the active document.


```vb
Sub ScratchPad() 
 
 Dim saPage As ScratchArea 
 Dim objFirst As Object 
 
 saPage = Application.ActiveDocument.ScratchArea 
 objFirst = saPage.Shapes(1) 
 
End Sub
```


