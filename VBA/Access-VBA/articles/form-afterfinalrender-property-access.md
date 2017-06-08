---
title: Form.AfterFinalRender Property (Access)
keywords: vbaac10.chm13548,vbaac10.chm5113
f1_keywords:
- vbaac10.chm13548,vbaac10.chm5113
ms.prod: access
api_name:
- Access.Form.AfterFinalRender
ms.assetid: c6e294f8-8cd9-1413-eff8-f2b033766326
ms.date: 06/08/2017
---


# Form.AfterFinalRender Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterFinalRender](form-afterfinalrender-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **AfterFinalRender**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeInsert** event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **AfterFinalRender** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterFinalRender = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

