---
title: Document.ShapesRemoved Event (Publisher)
keywords: vbapb10.chm285212677
f1_keywords:
- vbapb10.chm285212677
ms.prod: publisher
api_name:
- Publisher.Document.ShapesRemoved
ms.assetid: e2a67359-5673-2c72-e1fc-e3e3a3b564f9
ms.date: 06/08/2017
---


# Document.ShapesRemoved Event (Publisher)

Occurs when a shape is deleted from a publication.


## Syntax

 _expression_. **ShapesRemoved**

 _expression_A variable that represents a  **Document** object.


## Example

This example displays a message whenever a shape is removed from the active publication. For this example to work, you must place this code into the  **ThisDocument** module.


```vb
Private Sub Document_ShapesRemoved() 
 MsgBox "You just deleted one or more shapes." 
End Sub
```


