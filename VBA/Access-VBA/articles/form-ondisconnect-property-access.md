---
title: Form.OnDisconnect Property (Access)
keywords: vbaac10.chm13537,vbaac10.chm5101
f1_keywords:
- vbaac10.chm13537,vbaac10.chm5101
ms.prod: access
api_name:
- Access.Form.OnDisconnect
ms.assetid: 8f6514c7-8f61-2ae7-0859-8299523609ca
ms.date: 06/08/2017
---


# Form.OnDisconnect Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[OnDisconnect](form-ondisconnect-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **OnDisconnect**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **OnDisconnect** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).OnDisconnect = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

