---
title: Form.DataChange Property (Access)
keywords: vbaac10.chm13554
f1_keywords:
- vbaac10.chm13554
ms.prod: access
api_name:
- Access.Form.DataChange
ms.assetid: 14fd4c9c-eb18-8f4d-ebd9-6f389523c4cf
ms.date: 06/08/2017
---


# Form.DataChange Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[DataChange](form-datachange-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **DataChange**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **DataChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).DataChange = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

