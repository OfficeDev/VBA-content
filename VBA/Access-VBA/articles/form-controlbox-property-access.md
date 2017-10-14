---
title: Form.ControlBox Property (Access)
keywords: vbaac10.chm13372
f1_keywords:
- vbaac10.chm13372
ms.prod: access
api_name:
- Access.Form.ControlBox
ms.assetid: c4d9976c-631d-ae99-0c5d-e7008bbdadf9
ms.date: 06/08/2017
---


# Form.ControlBox Property (Access)

Specifies whether a form has a  **Control** menu in Form view and Datasheet view. Read/write **Boolean**.


## Syntax

 _expression_. **ControlBox**

 _expression_ A variable that represents a **Form** object.


## Remarks

The default value is  **True**.

Setting the  **ControlBox** property to **False** also removes the **Minimize**, **Maximize**, and **Close** buttons on a form.

It can only be set in form Design view.

To display a  **Control** menu on a form, the **ControlBox** property must be set to Yes and the form's **BorderStyle** property must be set to Thin, Sizable, or Dialog.

Even when a form's  **ControlBox** property is set to No, the form always has a **Control** menu when opened in Design view.

Setting the  **ControlBox** property to No suppresses the **Control** menu when you:


- Open the form in Form view from the Database window.
    
- Open the form from a macro.
    
- Open the form from Visual Basic.
    
- Open the form in Datasheet view.
    
- Switch to Form or Datasheet view from Design view.
    

## See also


#### Concepts


[Form Object](form-object-access.md)

