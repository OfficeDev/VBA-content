---
title: Screen.ActiveControl Property (Access)
keywords: vbaac10.chm12488
f1_keywords:
- vbaac10.chm12488
ms.prod: access
api_name:
- Access.Screen.ActiveControl
ms.assetid: 01d76377-c88d-8f64-b13b-c80f4d296834
ms.date: 06/08/2017
---


# Screen.ActiveControl Property (Access)

You can use the  **ActiveControl** property together with the **[Screen](screen-object-access.md)** object to identify or refer to the control that has the focus. Read-only **Control** object.


## Syntax

 _expression_. **ActiveControl**

 _expression_ A variable that represents a **Screen** object.


## Remarks

This property setting contains a reference to the  **Control** object that has the focus at run time.

You can use the  **ActiveControl** property to refer to the control that has the focus at run time together with one of its properties or methods. The following example assigns the name of the control with the focus to the `strControlName` variable.




```vb
Dim ctlCurrentControl As Control 
Dim strControlName As String 
Set ctlCurrentControl = Screen.ActiveControl 
strControlName = ctlCurrentControl.Name
```

If no control has the focus when you use the  **ActiveControl** property, or if all of the active form's controls are hidden or disabled, an error occurs.


## Example

The following example assigns the active control to the  `ctlCurrentControl` variable and then takes different actions depending on the value of the control's **Name** property.


```vb
Dim ctlCurrentControl As Control 
 
Set ctlCurrentControl = Screen.ActiveControl 
If ctlCurrentControl.Name = "txtCustomerID" Then 
 . 
 . ' Do something here. 
 . 
ElseIf ctlCurrentControl.Name = "btnCustomerDetails" Then 
 . 
 . ' Do something here. 
 . 
End If
```


## See also


#### Concepts


[Screen Object](screen-object-access.md)

