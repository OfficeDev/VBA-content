---
title: Form.PivotTableChange Property (Access)
keywords: vbaac10.chm13538,vbaac10.chm5102
f1_keywords:
- vbaac10.chm13538,vbaac10.chm5102
ms.prod: access
api_name:
- Access.Form.PivotTableChange
ms.assetid: d8d6a7eb-2bc1-e441-95fe-aefaec7fde9d
ms.date: 06/08/2017
---


# Form.PivotTableChange Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **PivotTableChange** event occurs. Read/write.


## Syntax

 _expression_. **PivotTableChange**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the  **PivotTableChange** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).PivotTableChange = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

