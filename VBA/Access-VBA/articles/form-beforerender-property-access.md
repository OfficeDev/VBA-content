---
title: Form.BeforeRender Property (Access)
keywords: vbaac10.chm13551
f1_keywords:
- vbaac10.chm13551
ms.prod: access
api_name:
- Access.Form.BeforeRender
ms.assetid: f80035ac-4ce6-ac8a-203f-c36afab5cd01
ms.date: 06/08/2017
---


# Form.BeforeRender Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[BeforeRender](form-beforerender-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **BeforeRender**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **BeforeRender** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeRender = "[Event Procedure]"
```


## See also


#### Concepts


[Form Object](form-object-access.md)

