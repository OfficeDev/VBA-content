---
title: CommandBarControls.Add Method (Office)
keywords: vbaof11.chm4001
f1_keywords:
- vbaof11.chm4001
ms.prod: office
api_name:
- Office.CommandBarControls.Add
ms.assetid: 53e2b0b9-b11a-bf52-a1a3-523aae2c35d8
ms.date: 06/08/2017
---


# CommandBarControls.Add Method (Office)

Creates a new  **CommandBarControl** object and adds it to the collection of controls on the specified command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Add**( **_Type_**, **_Id_**, **_Parameter_**, **_Before_**, **_Temporary_** )

 _expression_ Required. A variable that represents a **[CommandBarControls](commandbarcontrols-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|The type of control to be added to the specified command bar. Can be one of the following  **MsoControl** constants: **msoControlButton**, **msoControlEdit**, **msoControlDropdown**, **msoControlComboBox**, or **msoControlPopup**.|
| _Id_|Optional|**Variant**|An integer that specifies a built-in control. If the value of this argument is 1, or if this argument is omitted, a blank custom control of the specified type will be added to the command bar.|
| _Parameter_|Optional|**Variant**|For built-in controls, this argument is used by the container application to run the command. For custom controls, you can use this argument to send information to Visual Basic procedures, or you can use it to store information about the control (similar to a second  **Tag** property value).|
| _Before_|Optional|**Variant**|A number that indicates the position of the new control on the command bar. The new control will be inserted before the control at this position. If this argument is omitted, the control is added at the end of the specified command bar.|
| _Temporary_|Optional|**Variant**|**True** to make the new control temporary. controls are automatically deleted when the container application is closed. The default value is **False**.|

## Example

This example creates a custom editing toolbar that contains buttons (controls) for cutting, copying, and pasting.


```
Dim customBar As CommandBar 
Dim newButton As CommandBarButton 
Set customBar = CommandBars.Add("Custom") 
Set newButton = customBar.Controls _ 
    .Add(msoControlButton, CommandBars("Edit") _ 
    .Controls("Cut").Id) 
Set newButton = customBar.Controls _ 
    .Add(msoControlButton, CommandBars("Edit") _ 
    .Controls("Copy").Id) 
Set newButton = customBar.Controls _ 
    .Add(msoControlButton, CommandBars("Edit") _ 
    .Controls("Paste").Id) 
customBar.Visible = True
```


## See also


#### Concepts


[CommandBarControls Object](commandbarcontrols-object-office.md)
#### Other resources


[CommandBarControls Object Members](commandbarcontrols-members-office.md)

