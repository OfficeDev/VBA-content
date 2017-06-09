---
title: OptionButton.AfterUpdate Property (Access)
keywords: vbaac10.chm10609
f1_keywords:
- vbaac10.chm10609
ms.prod: access
api_name:
- Access.OptionButton.AfterUpdate
ms.assetid: 02ca295b-ff5c-2f6d-12f0-ea0bc176947a
ms.date: 06/08/2017
---


# OptionButton.AfterUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **AfterUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **AfterUpdate**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **AfterUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.


## Example

The following example specifies that when the AfterUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).AfterUpdate = "[Event Procedure]" 

```


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

