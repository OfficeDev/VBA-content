---
title: Form.CommandBeforeExecute Property (Access)
keywords: vbaac10.chm13542
f1_keywords:
- vbaac10.chm13542
ms.prod: access
api_name:
- Access.Form.CommandBeforeExecute
ms.assetid: 574568fa-e488-6d4d-a42f-07eb7c7f9536
ms.date: 06/08/2017
---


# Form.CommandBeforeExecute Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[CommandBeforeExecute](form-commandbeforeexecute-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **CommandBeforeExecute**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **CommandBeforeExecute** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).CommandBeforeExecute = "[Event Procedure]"
```


## See also


#### Concepts


[Form Object](form-object-access.md)

