---
title: BoundObjectFrame.BeforeUpdate Property (Access)
keywords: vbaac10.chm10961
f1_keywords:
- vbaac10.chm10961
ms.prod: access
api_name:
- Access.BoundObjectFrame.BeforeUpdate
ms.assetid: 01ee3c67-76c6-b651-042b-a7aa59e7443e
ms.date: 06/08/2017
---


# BoundObjectFrame.BeforeUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **BeforeUpdate**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the BeforeUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeUpdate = "[Event Procedure]" 

```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

