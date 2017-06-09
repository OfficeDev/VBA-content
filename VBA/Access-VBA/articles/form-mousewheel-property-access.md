---
title: Form.MouseWheel Property (Access)
keywords: vbaac10.chm13552
f1_keywords:
- vbaac10.chm13552
ms.prod: access
api_name:
- Access.Form.MouseWheel
ms.assetid: 364f7854-d7d5-5fe2-effa-6154e86376b4
ms.date: 06/08/2017
---


# Form.MouseWheel Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **MouseWheel** event occurs. Read/write.


## Syntax

 _expression_. **MouseWheel**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **MouseWheel** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).MouseWheel = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

