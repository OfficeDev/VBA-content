---
title: Form.BeforeInsert Property (Access)
keywords: vbaac10.chm13432
f1_keywords:
- vbaac10.chm13432
ms.prod: access
api_name:
- Access.Form.BeforeInsert
ms.assetid: 634b0480-ddb3-7ef7-b347-57ca9a4eebad
ms.date: 06/08/2017
---


# Form.BeforeInsert Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeInsert](form-beforeinsert-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **BeforeInsert**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **BeforeInsert** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeInsert = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

