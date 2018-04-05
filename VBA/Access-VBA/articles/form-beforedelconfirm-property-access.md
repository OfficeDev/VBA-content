---
title: Form.BeforeDelConfirm Property (Access)
keywords: vbaac10.chm13438
f1_keywords:
- vbaac10.chm13438
ms.prod: access
api_name:
- Access.Form.BeforeDelConfirm
ms.assetid: 8926afb1-5a86-eddd-5b3f-68abe83fb076
ms.date: 06/08/2017
---


# Form.BeforeDelConfirm Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the[BeforeDelConfirm](form-beforedelconfirm-event-access.md)event occurs. Read/write.


## Syntax

 _expression_. **BeforeDelConfirm**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeDelConfirm event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the BeforeDelConfirm event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeDelConfirm = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

