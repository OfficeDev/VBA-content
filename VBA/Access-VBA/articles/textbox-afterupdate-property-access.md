---
title: TextBox.AfterUpdate Property (Access)
keywords: vbaac10.chm11116
f1_keywords:
- vbaac10.chm11116
ms.prod: access
api_name:
- Access.TextBox.AfterUpdate
ms.assetid: 690bc0cd-9717-7712-c022-75ba457ca0e3
ms.date: 06/08/2017
---


# TextBox.AfterUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **AfterUpdate**

 _expression_ A variable that represents a **TextBox** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **AfterUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the AfterUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterUpdate = "[Event Procedure]" 

```


## See also


#### Concepts


[TextBox Object](textbox-object-access.md)

