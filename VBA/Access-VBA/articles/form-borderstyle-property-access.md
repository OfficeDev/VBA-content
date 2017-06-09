---
title: Form.BorderStyle Property (Access)
keywords: vbaac10.chm13371
f1_keywords:
- vbaac10.chm13371
ms.prod: access
api_name:
- Access.Form.BorderStyle
ms.assetid: a6c4d49b-4227-09e9-2999-6f8954bbeb39
ms.date: 06/08/2017
---


# Form.BorderStyle Property (Access)

Specifies the type of border and border elements (title bar,  **Control** menu, **Minimize** and **Maximize** buttons, or **Close** button) to use for the form. You typically use different border styles for normal forms, pop-up forms, and custom dialog boxes. Read/write **Byte**.


## Syntax

 _expression_. **BorderStyle**

 _expression_ A variable that represents a **Form** object.


## Remarks

For controls, the  **BorderStyle** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|None|0|The form has no border or related border elements. The form isn't resizable.|
|Thin|1|The form has a thin border and can include any of the border elements. The form isn't resizable (the  **Size** command on the **Control** menu isn't available). You often use this setting for pop-up forms. (If you want a form to remain on top of all Microsoft Access windows, you must also set its **PopUp** property to Yes.)|
|Sizable|2|(Default) The form has the default border for Microsoft Access forms, can include any of the border elements, and can be resized. You often use this setting for normal Microsoft Access forms.|
|Dialog|3|The form has a thick (double) border and can include only a title bar,  **Close** button, and **Control** menu. The form can't be maximized, minimized, or resized (the **Maximize**, **Minimize**, and **Size** commands aren't available on the **Control** menu). You often use this setting for custom dialog boxes. (If you want a form to be modal, however, you must also set its **Modal** property to Yes. If you want it to be a modal pop-up form, like most dialog boxes, you must set both its **PopUp** and **Modal** properties to Yes.)|
For a form, the  **BorderStyle** property establishes the characteristics that visually identify the form as a normal form, a pop-up form, or a custom dialog box. You may also set the **Modal** and **PopUp** properties to further define the form's characteristics.

You may also want to set the form's  **ControlBox**, **CloseButton**, **MinMaxButtons**, **ScrollBars**, **NavigationButtons**, and **RecordSelectors** properties. These properties interact in the following ways:


- If the  **BorderStyle** property is set to None or Dialog, the form doesn't have **Maximize** or **Minimize** buttons, regardless of its **MinMaxButtons** property setting.
    
- If the  **BorderStyle** property is set to None, the form doesn't have a **Control** menu, regardless of its **ControlBox** property setting.
    
- The  **BorderStyle** property setting doesn't affect the display of the scroll bars, navigation buttons, record number box, or record selectors.
    
The  **BorderStyle** property takes effect only in Form view. The property setting is ignored in form Design view.

If you set the  **BorderStyle** property of a pop-up form to None, you won't be able to close the form unless you add a **Close** button to it that runs a macro containing the **Close** action or an event procedure in Visual Basic that uses the **Close** method.

Pop-up forms are typically fixed in size, but you can make a pop-up form sizable by setting its  **PopUp** property to Yes and its **BorderStyle** property to Sizable.


## See also


#### Concepts


[Form Object](form-object-access.md)

