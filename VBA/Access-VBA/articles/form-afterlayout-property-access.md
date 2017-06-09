---
title: Form.AfterLayout Property (Access)
keywords: vbaac10.chm13550
f1_keywords:
- vbaac10.chm13550
ms.prod: access
api_name:
- Access.Form.AfterLayout
ms.assetid: 8d548e7b-6d68-4631-2c59-f6b8d39cbb12
ms.date: 06/08/2017
---


# Form.AfterLayout Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the[AfterLayout](form-afterlayout-event-access.md)event occurs. Read/write.


## Syntax

 _expression_. **AfterLayout**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **AfterLayout** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterLayout = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

