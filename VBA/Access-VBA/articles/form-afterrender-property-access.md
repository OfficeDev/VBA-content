---
title: Form.AfterRender Property (Access)
keywords: vbaac10.chm13549,vbaac10.chm5114
f1_keywords:
- vbaac10.chm13549,vbaac10.chm5114
ms.prod: access
api_name:
- Access.Form.AfterRender
ms.assetid: 868b9a9d-a1e3-d460-fa7c-26cb5791c5ad
ms.date: 06/08/2017
---


# Form.AfterRender Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterRender](form-afterrender-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **AfterRender**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **AfterRender** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterRender = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

