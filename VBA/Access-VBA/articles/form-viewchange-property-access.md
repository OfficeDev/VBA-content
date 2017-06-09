---
title: Form.ViewChange Property (Access)
keywords: vbaac10.chm13553,vbaac10.chm5118
f1_keywords:
- vbaac10.chm13553,vbaac10.chm5118
ms.prod: access
api_name:
- Access.Form.ViewChange
ms.assetid: f8a8fe82-6983-5632-b779-879faf228ac2
ms.date: 06/08/2017
---


# Form.ViewChange Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[ViewChange](form-viewchange-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **ViewChange**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **ViewChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).ViewChange = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

