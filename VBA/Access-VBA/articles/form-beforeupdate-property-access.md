---
title: Form.BeforeUpdate Property (Access)
keywords: vbaac10.chm13434
f1_keywords:
- vbaac10.chm13434
ms.prod: access
api_name:
- Access.Form.BeforeUpdate
ms.assetid: b4b39ab8-e37c-8803-b6c3-032707342c92
ms.date: 06/08/2017
---


# Form.BeforeUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **BeforeUpdate**

 _expression_ A variable that represents a **Form** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the BeforeUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeUpdate = "[Event Procedure]" 

```


## See also


#### Concepts


[Form Object](form-object-access.md)

