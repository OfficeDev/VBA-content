---
title: Form.BeforeQuery Property (Access)
keywords: vbaac10.chm13540
f1_keywords:
- vbaac10.chm13540
ms.prod: access
api_name:
- Access.Form.BeforeQuery
ms.assetid: 40e763fd-897a-a0b1-72a9-d73ec628e397
ms.date: 06/08/2017
---


# Form.BeforeQuery Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeQuery](form-beforequery-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **BeforeQuery**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **BeforeQuery** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0)
```


## See also


#### Concepts


[Form Object](form-object-access.md)

