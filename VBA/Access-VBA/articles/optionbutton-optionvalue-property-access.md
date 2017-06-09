---
title: OptionButton.OptionValue Property (Access)
keywords: vbaac10.chm10568
f1_keywords:
- vbaac10.chm10568
ms.prod: access
api_name:
- Access.OptionButton.OptionValue
ms.assetid: 23e170c7-21ac-4725-b54b-ad778bba9f31
ms.date: 06/08/2017
---


# OptionButton.OptionValue Property (Access)

Each control in an option group has a numeric value that you can set with the  **OptionValue** property. Read/write **Long**.


## Syntax

 _expression_. **OptionValue**

 _expression_ A variable that represents an **OptionButton** object.


## Remarks

When the control is selected, the number is assigned to the option group. If the option group is bound to a field, the value of the selected control's  **OptionValue** property is stored in the field.

For example, this  **Region** option group is bound to the Region field in a table. The **Europe** button has an **OptionValue** property setting of 1, the **Asia** button has a setting of 2, and the **Africa** button has a setting of 3. When one of these buttons is selected, the **Region** option group value will be the same as the **OptionValue** property setting for the selected control. In this case, because the **Region** option group is bound to the Region field, the value of this field in the table also equals 2.

The  **OptionValue** property applies only to the check box, option button, and toggle button controls in an option group.

Unless you change the  **OptionValue** property yourself, the first control you place in an option group has a value of 1, the second control has a value of 2, and so on.

The  **OptionValue** property is only available when the control is placed inside an option group control. When a check box, a toggle button, or an option button isn't in an option group, the control has no **OptionValue** property. Instead, each such control has a **ControlSource** property, and the value of each control will be either **True** if selected or **False** if not selected.


## Example

The following example sets the  **OptionValue** property for three option buttons in the "Ship Method Group" option group when a form opens. When an option button is selected in the option group, a message displays indicating the shipper's assigned ID number.


```vb
Private Sub Form_Open(Cancel As Integer) 
 
 Me.Controls("ABC Couriers").OptionValue = 15876 
 Me.Controls("Speedy Delivery").OptionValue = 742 
 Me.Controls("Lightning Express").OptionValue = 1256 
 
End Sub 
 
Private Sub Ship_Method_Group_Click() 
 
 MsgBox "The ID for the selected shipper is " &; Me.Controls("Ship Method Group").Value 
 
End Sub
```


## See also


#### Concepts


[OptionButton Object](optionbutton-object-access.md)

