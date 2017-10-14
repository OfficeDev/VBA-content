---
title: Form.CommandChecked Property (Access)
keywords: vbaac10.chm13543
f1_keywords:
- vbaac10.chm13543
ms.prod: access
api_name:
- Access.Form.CommandChecked
ms.assetid: 4f3bb0fa-6f3f-4836-a0d0-06d480e1d194
ms.date: 06/08/2017
---


# Form.CommandChecked Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandChecked](form-commandchecked-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **CommandChecked**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **CommandChecked** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).CommandChecked = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

