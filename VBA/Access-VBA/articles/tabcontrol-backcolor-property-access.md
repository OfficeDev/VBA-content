---
title: TabControl.BackColor Property (Access)
keywords: vbaac10.chm10839
f1_keywords:
- vbaac10.chm10839
ms.prod: access
api_name:
- Access.TabControl.BackColor
ms.assetid: 801d889c-0741-1de5-48ed-ea1db8f9a75b
ms.date: 06/08/2017
---


# TabControl.BackColor Property (Access)

Gets or sets the interior color of the specified object. Read/write  **Long**.


## Syntax

 _expression_. **BackColor**

 _expression_ A variable that represents a **TabControl** object.


## Remarks

The  **BackColor** property contains a numeric expression that corresponds to the color used to fill a control's or section's interior.

You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

To use the  **BackColor** property, the **BackStyle** property, if available, must be set to Normal.


## Example

The following example uses the  **RGB** function to set the , **BackColor**, and **ForeColor** properties depending on the value of the `txtPastDue` text box. You can also use the **QBColor** function to set these properties. Putting the following code in the Form_Current( ) event sets the control display characteristics as soon as the user opens a form or moves to a new record.


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


[TabControl Object](tabcontrol-object-access.md)

