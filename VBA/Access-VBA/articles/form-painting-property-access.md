---
title: Form.Painting Property (Access)
keywords: vbaac10.chm13416
f1_keywords:
- vbaac10.chm13416
ms.prod: access
api_name:
- Access.Form.Painting
ms.assetid: 6fbbd097-8882-b633-bbd6-9dcc0bb31db9
ms.date: 06/08/2017
---


# Form.Painting Property (Access)

You can use the  **Painting** property to specify whether a form is repainted. Read/write **Boolean**.


## Syntax

 _expression_. **Painting**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property can be set and applies only in Form view and is unavailable in other views.

The  **Painting** property is similar to the Echo action. However, the **Painting** property prevents repainting of a single form, whereas the Echo action prevents repainting of all open windows in an application.

Setting the  **Painting** property for a form to **False** also prevents all controls (except subform controls) on a form from being repainted. To prevent a subform control from being repainted, you must set the **Painting** property for the subform to **False**. (Note that you set the **Painting** property for the subform, not the subform control.)

The  **Painting** property is automatically set to **True** whenever the form gets or loses the focus. You can set this property to **False** while you are working on a form if you don't want to see changes to the form or to its controls. For example, if a form has a set of controls that are automatically resized when the form is resized and you don't want the user to see each individual control move, you can turn **Painting** off, and then move all of the controls, then turn **Painting** back on.


## Example

The following example uses the  **Painting** property to enable or disable form painting depending on whether the `SetPainting` variable is set to **True** or **False**. If form painting is turned off, Microsoft Access displays the hourglass icon while painting is turned off.


```vb
Public Sub EnablePaint(ByRef frmName As Form, _ 
 ByVal SetPainting As Integer) 
 
 frmName.Painting = SetPainting 
 
 ' Form painting is turned off. 
 If SetPainting = False Then 
 DoCmd.Hourglass True 
 Else 
 DoCmd.Hourglass False 
 End If 
 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

