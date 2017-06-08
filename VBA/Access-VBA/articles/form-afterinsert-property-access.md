---
title: Form.AfterInsert Property (Access)
keywords: vbaac10.chm13433
f1_keywords:
- vbaac10.chm13433
ms.prod: access
api_name:
- Access.Form.AfterInsert
ms.assetid: 95bc1f0d-a0fa-ffdd-ef5a-e6eb2a854feb
ms.date: 06/08/2017
---


# Form.AfterInsert Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterInsert](form-afterinsert-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **AfterInsert**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeInsert** event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **AfterInsert** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).After Insert = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

