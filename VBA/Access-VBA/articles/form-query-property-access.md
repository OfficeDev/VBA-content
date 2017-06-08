---
title: Form.Query Property (Access)
keywords: vbaac10.chm13539,vbaac10.chm5103
f1_keywords:
- vbaac10.chm13539,vbaac10.chm5103
ms.prod: access
api_name:
- Access.Form.Query
ms.assetid: fcef59f9-f405-0a05-f986-b29c2b0528de
ms.date: 06/08/2017
---


# Form.Query Property (Access)

Returns or sets a  **String** indicating which macro, event procedure, or user-defined function runs when the **[Query](form-query-event-access.md)** event occurs. Read/write.


## Syntax

 _expression_. **Query**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are "macroname" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the BeforeInsert event for the specified object, or "=functionname()" where _functionname_ is the name of a user-defined function.

This property corresponds to the  **On Query** property that is displayed in a form's property sheet.


## Example

The following example specifies that when the  **Query** event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).Query = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

