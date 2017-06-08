---
title: Form.OnConnect Property (Access)
keywords: vbaac10.chm13536,vbaac10.chm5100
f1_keywords:
- vbaac10.chm13536,vbaac10.chm5100
ms.prod: access
api_name:
- Access.Form.OnConnect
ms.assetid: de181e49-ccba-52fa-f521-3e55f3ed78d2
ms.date: 06/08/2017
---


# Form.OnConnect Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[OnConnect](form-onconnect-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **OnConnect**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **OnConnect** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).OnConnect = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

