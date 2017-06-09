---
title: Form.PopUp Property (Access)
keywords: vbaac10.chm13369
f1_keywords:
- vbaac10.chm13369
ms.prod: access
api_name:
- Access.Form.PopUp
ms.assetid: 0ccaa174-80e2-5ca3-9614-93b12dc1bfcd
ms.date: 06/08/2017
---


# Form.PopUp Property (Access)

Specifies whether a form opens as a pop-up window. Read/write  **Boolean**.


## Syntax

 _expression_. **PopUp**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **PopUp** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|**True**|The form opens as a pop-up window. It remains on top of all other Microsoft Access windows.|
|No|**False**|(Default) The form isn't a pop-up window.|
The  **PopUp** property can be set only in form Design view.

To specify the type of border you want on a pop-up window, use the  **BorderStyle** property. You typically set the **BorderStyle** property to Thin for pop-up windows.

To create a custom dialog box, set the  **Modal** property to Yes, the **PopUp** property to Yes, and the **BorderStyle** property to Dialog.

Setting the  **PopUp** property to Yes makes the form a pop-up window only when you do one of the following:


- Open it in Form view from the Database window.
    
- Open it in Form view by using a macro or Visual Basic.
    
- Switch from Design view to Form view.
    
When the  **PopUp** property is set to Yes, you can't switch to other views from Form view because the form's toolbar isn't available. (You can't switch a pop-up form from Form view to Datasheet view, even in a macro or Visual Basic.) You must close the form and reopen it in Design or Datasheet view.

The form isn't a pop-up form in Design or Datasheet view, and also isn't if you switch from Datasheet to Form view.


 **Note**  You can use the Dialog setting of the Window Mode argument of the OpenForm action to open a form with its  **PopUp** and **Modal** properties set to Yes.

When you maximize a window in Microsoft Access, all other windows are also maximized when you open them or switch to them. However, pop-up forms aren't maximized. If you want a form to maintain its size when other windows are maximized, set its  **PopUp** property to Yes.


## See also


#### Concepts


[Form Object](form-object-access.md)

