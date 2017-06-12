---
title: CommandButton.ForeColor Property (Access)
keywords: vbaac10.chm10471
f1_keywords:
- vbaac10.chm10471
ms.prod: access
api_name:
- Access.CommandButton.ForeColor
ms.assetid: 6d19e4b2-2375-fe37-c226-4489ebcb808e
ms.date: 06/08/2017
---


# CommandButton.ForeColor Property (Access)

You can use the  **ForeColor** property to specify the color for text in a control. Read/write **Long**.


## Syntax

 _expression_. **ForeColor**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

You can use this property for controls on forms or reports to make them easy to read or to convey a special meaning. For example, you can change the color of the text in the UnitsInStock control when its value falls below the reorder level.

You can also use this property on reports to create special visual effects when you print with a color printer. When used on a report, this property specifies the printing and drawing color for the  **[Print](report-print-method-access.md)**, **[Line](report-line-method-access.md)**, and **[Circle](report-circle-method-access.md)** methods.

The  **ForeColor** property contains a numeric expression that represents the value of the text color in the control.

You can use the Color Builder to set this property by clicking the  **Build** button to the right of the property box in the property sheet. Using the Color Builder enables you to define custom colors for text in controls.

You can set the default for this property by using a control's default control style or the  **DefaultControl** property in Visual Basic.

For reports, you can set the  **Circle** property only by using a macro or a Visual Basic event procedure specified in a section's **OnPrint** event property setting.


## Example

The following example uses the  **RGB** function to set the **BorderColor**, **BackColor**, and **ForeColor** properties depending on the value of the `txtPastDue` text box. You can also use the **QBColor** function to set these properties. Putting the following code in the Form_Current( ) event sets the control display characteristics as soon as the user opens a form or moves to a new record.


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


[CommandButton Object](commandbutton-object-access.md)

