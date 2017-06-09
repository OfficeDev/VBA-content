---
title: ComboBox.BeforeUpdate Property (Access)
keywords: vbaac10.chm11447
f1_keywords:
- vbaac10.chm11447
ms.prod: access
api_name:
- Access.ComboBox.BeforeUpdate
ms.assetid: ce748fb1-4f8d-9e96-f77c-5dfc54dfee48
ms.date: 06/08/2017
---


# ComboBox.BeforeUpdate Property (Access)

Returns or sets which macro, event procedure, or user-defined function runs when the  **BeforeUpdate** event occurs. Read/write **String**.


## Syntax

 _expression_. **BeforeUpdate**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

Valid values for this property are " _macroname_" where  _macroname_ is the name of macro, "[Event Procedure]" which indicates the event procedure associated with the **BeforeUpdate** event for the specified object, or " **=** _functionname_ **()** " where _functionname_ is the name of a user-defined function.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) Luke Chung,[FMS, Inc.](http://www.fmsinc.com/)


- [Tips and Techniques for Using and Validating Combo Boxes](http://www.fmsinc.com/free/NewTips/Access/ComboBox/AccessComboBox.asp)
    

## Example

The following example specifies that when the BeforeUpdate event occurs on the first form of the current project, the associated event procedure should run.


```vb
Forms(0).BeforeUpdate = "[Event Procedure]" 

```


## About the Contributors
<a name="AboutContributors"> </a>

Luke Chung is the founder and president of FMS, Inc., a leading provider of custom database solutions and developer tools. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[ComboBox Object](combobox-object-access.md)

