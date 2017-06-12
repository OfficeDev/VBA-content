---
title: BoundObjectFrame.BorderColor Property (Access)
keywords: vbaac10.chm10936
f1_keywords:
- vbaac10.chm10936
ms.prod: access
api_name:
- Access.BoundObjectFrame.BorderColor
ms.assetid: e3c43808-1254-2635-264e-2f3e79cb2c8a
ms.date: 06/08/2017
---


# BoundObjectFrame.BorderColor Property (Access)

You can use the  **BorderColor** property to specify the color of a control's border. Read/write **Long**.


## Syntax

 _expression_. **BorderColor**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

The  **BorderColor** property setting is a numeric expression that corresponds to the color you want to use for a control's border.

You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

A control's border color is visible only when its  **SpecialEffect** property is set to Flat or Shadowed. If the **SpecialEffect** property is set to something other than Flat or Shadowed, setting the **BorderColor** property changes the **SpecialEffect** property setting to Flat.


## Example

The following example uses the  **RGB** function to set the **BorderColor**, **BackColor**, and **ForeColor** properties depending on the value of the `txtPastDue` text box. You can also use the QBColor function to set these properties. Putting the following code in the Form_Current( ) event sets the control display characteristics as soon as the user opens a form or moves to a new record.


```vb
Sub Form_Current() 
 Dim curAmntDue As Currency, lngBlack As Long 
 Dim lngRed As Long, lngYellow As Long, lngWhite As Long 
 
 If Not IsNull(Me!txtPastDue.Value) Then 
 curAmntDue = Me!txtPastDue.Value 
 Else 
 Exit Sub 
 End If 
 lngRed = RGB(255, 0, 0) 
 lngBlack = RGB(0, 0, 0) 
 lngYellow = RGB(255, 255, 0) 
 lngWhite = RGB(255, 255, 255) 
 If curAmntDue > 100 Then 
 Me!txtPastDue.BorderColor = lngRed 
 Me!txtPastDue.ForeColor = lngRed 
 Me!txtPastDue.BackColor = lngYellow 
 Else 
 Me!txtPastDue.BorderColor = lngBlack 
 Me!txtPastDue.ForeColor = lngBlack 
 Me!txtPastDue.BackColor = lngWhite 
 End If 
End Sub
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

