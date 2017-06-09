---
title: Form.SelectionChange Property (Access)
keywords: vbaac10.chm13541
f1_keywords:
- vbaac10.chm13541
ms.prod: access
api_name:
- Access.Form.SelectionChange
ms.assetid: e31876fc-103a-d231-a6fa-7cb026a343e1
ms.date: 06/08/2017
---


# Form.SelectionChange Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[SelectionChange](form-selectionchange-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **SelectionChange**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **SelectionChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).SelectionChange = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

