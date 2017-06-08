---
title: CommandButton.Default Property (Access)
keywords: vbaac10.chm10455
f1_keywords:
- vbaac10.chm10455
ms.prod: access
api_name:
- Access.CommandButton.Default
ms.assetid: b643350e-9a89-a0ff-b8dd-f1c2c1392992
ms.date: 06/08/2017
---


# CommandButton.Default Property (Access)

You can use the  **Default** property to specify whether a command button is the default button on a form. Read/write **Boolean**.


## Syntax

 _expression_. **Default**

 _expression_ A variable that represents a **CommandButton** object.


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


[CommandButton Object](commandbutton-object-access.md)

