---
title: Form.Modal Property (Access)
keywords: vbaac10.chm13370
f1_keywords:
- vbaac10.chm13370
ms.prod: access
api_name:
- Access.Form.Modal
ms.assetid: a36b42f6-9d97-acea-cda3-2f380a3270c2
ms.date: 06/08/2017
---


# Form.Modal Property (Access)

You can use the  **Modal** property to specify whether a form opens as a modal window. When a form opens as a modal window, you must close the window before you can move the focus to another object. Read/write **Boolean**.


## Syntax

 _expression_. **Modal**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **Modal** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The form or report opens as a modal window.|
|No|**False**|(Default) The form opens as a non-modal window.|
When you open a modal window, other windows in Microsoft Access are disabled until you close the form (although you can switch to windows in other applications). To disable menus and toolbars in addition to other windows, set both the  **Modal** and **PopUp** properties to Yes.

You can use the  **BorderStyle** property to specify the kind of border a form will have. Typically, modal forms have the **BorderStyle** property set to Dialog.

You can use the  **Modal**, **PopUp**, and **BorderStyle** properties to create a custom dialog box. You can set **Modal** to Yes, **PopUp** to Yes, and **BorderStyle** to Dialog for custom dialog boxes.

Setting the  **Modal** property to Yes makes the form modal only when you:


- Open it in Form view from the Database window.
    
- Open it in Form view by using a macro or Visual Basic.
    
- Switch from Design view to Form view.
    
When the form is modal, you can't switch to Datasheet view from Form view, although you can switch to Design view and then to Datasheet view.

The form isn't modal in Design view or Datasheet view and also isn't modal if you switch from Datasheet view to Form view.


 **Note**  You can use the Dialog setting of the Window Mode action argument of the OpenForm action to open a form with its  **Modal** and **PopUp** properties set to Yes.


## See also


#### Concepts


[Form Object](form-object-access.md)

