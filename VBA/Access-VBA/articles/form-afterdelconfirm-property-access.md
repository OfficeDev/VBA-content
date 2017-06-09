---
title: Form.AfterDelConfirm Property (Access)
keywords: vbaac10.chm13439,vbaac10.chm4085
f1_keywords:
- vbaac10.chm13439,vbaac10.chm4085
ms.prod: access
api_name:
- Access.Form.AfterDelConfirm
ms.assetid: fcc1585b-ddb9-7b39-aa21-07de0e50ac00
ms.date: 06/08/2017
---


# Form.AfterDelConfirm Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[AfterDelConfirm](form-afterdelconfirm-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **AfterDelConfirm**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeInsert** event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **AfterDelConfirm** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).After DelConfirm = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

