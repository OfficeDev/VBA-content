---
title: CustomControl.Default Property (Access)
keywords: vbaac10.chm12045
f1_keywords:
- vbaac10.chm12045
ms.prod: access
api_name:
- Access.CustomControl.Default
ms.assetid: ffe92e84-4bfa-56a2-298e-00d448f8dc29
ms.date: 06/08/2017
---


# CustomControl.Default Property (Access)

You can use the  **Default** property to specify whether a command button is the default button on a form. Read/write **Boolean**.


## Syntax

 _expression_. **Default**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

The  **Default** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The command button is the default button.|
|No|**False**|(Default) The command button isn't the default button.|
When the command button's  **Default** property setting is Yes and the Form window is active, the user can choose the command button by pressing ENTER (if no other command button has the focus) as well as by clicking the command button.

Only one command button on a form can be the default button. When the  **Default** property is set to Yes for one command button, it is automatically set to No for all other command buttons on the form.

For a form that supports irreversible operations, such as deletions, it's a good idea to make the  **Cancel** button the default command button. To do this, set both the **Default** property and the Cancel property to Yes.


## See also


#### Concepts


[CustomControl Object](customcontrol-object-access.md)

