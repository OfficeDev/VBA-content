---
title: Document.UndoClear Method (Publisher)
keywords: vbapb10.chm196705
f1_keywords:
- vbapb10.chm196705
ms.prod: publisher
api_name:
- Publisher.Document.UndoClear
ms.assetid: 63e9bb00-950f-3e30-3897-434362b9efbf
ms.date: 06/08/2017
---


# Document.UndoClear Method (Publisher)

Clears the list of actions that can be undone for the specified publication. Corresponds to the list of items that appears when you click the arrow beside the  **Undo** button on the **Standard** toolbar.


## Syntax

 _expression_. **UndoClear**

 _expression_A variable that represents a  **Document** object.


## Remarks

Include this method at the end of a macro to keep Microsoft Visual Basic actions from appearing in the  **Undo** box (for example, "VBA-Selection.InsertAfter").


## Example

This example clears the list of actions that can be undone for the active publication.


```vb
ActiveDocument.UndoClear
```


