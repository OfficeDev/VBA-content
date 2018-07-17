---
title: Screen.ActiveForm Property (Access)
keywords: vbaac10.chm12490
f1_keywords:
- vbaac10.chm12490
ms.prod: access
api_name:
- Access.Screen.ActiveForm
ms.assetid: 5cf41661-656e-e62f-530e-0d2fa5466146
ms.date: 06/08/2017
---


# Screen.ActiveForm Property (Access)

You can use the  **ActiveForm** property together with the **[Screen](screen-object-access.md)** object to identify or refer to the form that has the focus. Read-only **Form** object.


## Syntax

 _expression_. **ActiveForm**

 _expression_ A variable that represents a **Screen** object.


## Remarks

This property setting contains a reference to the  **[Form](form-object-access.md)** object that has the focus at run time.

You can use the  **ActiveForm** property to refer to an active form together with one of its properties or methods. The following example displays the **Name** property setting of the active form.




```vb
Dim frmCurrentForm As Form 
Set frmCurrentForm = Screen.ActiveForm 
MsgBox "Current form is " &; frmCurrentForm.Name
```

If a subform has the focus,  **ActiveForm** refers to the main form. If no form or subform has the focus when you use the **ActiveForm** property, an error occurs.


## See also


#### Concepts


[Screen Object](screen-object-access.md)

