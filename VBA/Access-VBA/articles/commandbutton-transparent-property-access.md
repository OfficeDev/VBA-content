---
title: CommandButton.Transparent Property (Access)
keywords: vbaac10.chm10454,vbaac10.chm4526
f1_keywords:
- vbaac10.chm10454,vbaac10.chm4526
ms.prod: access
api_name:
- Access.CommandButton.Transparent
ms.assetid: 655e127e-7e2e-c2c2-a979-952f95c534a6
ms.date: 06/08/2017
---


# CommandButton.Transparent Property (Access)

You can use the  **Transparent** property to specify whether a command button is solid or transparent. Read/write **Boolean**.


## Syntax

 _expression_. **Transparent**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

Use the  **BackStyle** property to make other controls solid or transparent.

You can use this property to place a transparent command button over another control. For example, you could place several transparent buttons over a picture displayed in an image control and run various macros or Visual Basic event procedures depending on which part of the picture the user clicks.


 **Note**  To hide and disable a button, use the  **Visible** property. To disable a button without hiding it, use the **Enabled** property. To hide a button only when a form or report is printed, use the **DisplayWhen** property.


## Example

The following example makes the command button "Preview" on the "Purchase Orders" form transparent.


```vb
Forms.Item("Purchase Orders").Controls.Item("Preview"). _ 
 Transparent = True
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

