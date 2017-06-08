---
title: OptionGroup.BeforeUpdate Property (Access)
keywords: vbaac10.chm10862
f1_keywords:
- vbaac10.chm10862
ms.prod: access
api_name:
- Access.OptionGroup.BeforeUpdate
ms.assetid: 0ea86e13-03ba-9f56-ef42-e8147fa70064
ms.date: 06/08/2017
---


# OptionGroup.BeforeUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **BeforeUpdate**

 _expression_ A variable that represents an **OptionGroup** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the BeforeUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeUpdate = "[Event Procedure]" 

```


## See also


#### Concepts


[OptionGroup Object](optiongroup-object-access.md)

