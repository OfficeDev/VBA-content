---
title: Document.ShapesAdded Event (Publisher)
keywords: vbapb10.chm285212675
f1_keywords:
- vbapb10.chm285212675
ms.prod: publisher
api_name:
- Publisher.Document.ShapesAdded
ms.assetid: f6573f7c-56fa-1efa-9dba-39cde3859cc0
ms.date: 06/08/2017
---


# Document.ShapesAdded Event (Publisher)

Occurs when one or more new shapes are added to a publication. This event occurs whether shapes are added manually or programmatically.


## Syntax

 _expression_. **ShapesAdded**

 _expression_A variable that represents a  **Document** object.


## Example

This example displays a message whenever a new shape is added to the active publication. For this example to work, you must place this code into the  **ThisDocument** module.


```vb
Private Sub PubDoc_ShapesAdded() 
 MsgBox "You just added a new shape." 
End Sub
```


