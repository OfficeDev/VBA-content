---
title: CommandBarComboBox.Id Property (Office)
ms.prod: office
api_name:
- Office.CommandBarComboBox.Id
ms.assetid: 9cc143cb-4063-b397-05c9-d50a7c2efcb0
ms.date: 06/08/2017
---


# CommandBarComboBox.Id Property (Office)

Gets the ID for a built-in  **CommandBarComboBox** control. Read-only.


## Syntax

 _expression_. **Id**

 _expression_ Required. A variable that represents a **[CommandBarComboBox](commandbarcombobox-object-office.md)** object.


## Remarks

A control's ID determines the built-in action for that control. The value of the  **Id** property for all custom controls is 1.


## Example

This example changes the button face of the first control on the command bar named "Custom2" if the button's  **ID** value is less than 25.


```
Set ctrl = CommandBars("Custom").Controls(1) 
With ctrl 
 If .Id < 25 Then 
 .FaceId = 17 
 .Tag = "Changed control" 
 End If 
End With
```

The following example changes the caption of every control on the toolbar named "Standard" to the current value of the  **Id** property for that control.




```
For Each ctl In CommandBars("Standard").Controls 
 ctl.Caption = CStr(ctl.Id) 
Next ctl
```


## See also


#### Concepts


[CommandBarComboBox Object](commandbarcombobox-object-office.md)
#### Other resources


[CommandBarComboBox Object Members](commandbarcombobox-members-office.md)

