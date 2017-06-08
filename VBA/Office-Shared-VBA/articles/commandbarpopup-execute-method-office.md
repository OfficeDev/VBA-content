---
title: CommandBarPopup.Execute Method (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Execute
ms.assetid: fedebe76-86f5-9c30-6e23-a20e0024bbf4
ms.date: 06/08/2017
---


# CommandBarPopup.Execute Method (Office)

Runs the procedure or built-in command assigned to the specified  **CommandBarPopup** control.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Execute**

 _expression_ Required. A variable that represents a **[CommandBarPopup](commandbarpopup-object-office.md)** object.


## Example

This Microsoft Excel example creates a command bar and then adds a built-in command bar button control to it. The button executes the Excel  **AutoSum** function. This example uses the **Execute** method to total the selected range of cells when the command bar appears.


```
Dim cbrCustBar As CommandBar 
Dim ctlAutoSum As CommandBarButton 
Set cbrCustBar = CommandBars.Add("Custom") 
Set ctlAutoSum = cbrCustBar.Controls _ 
    .Add(msoControlButton, CommandBars("Standard") _ 
    .Controls("AutoSum").Id) 
cbrCustBar.Visible = True  
ctlAutoSum.Execute
```


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

