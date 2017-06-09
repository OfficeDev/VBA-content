---
title: CommandBarControl.Execute Method (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.Execute
ms.assetid: 5b95846f-99c6-93b3-2167-6bd7acf5d508
ms.date: 06/08/2017
---


# CommandBarControl.Execute Method (Office)

Runs the procedure or built-in command assigned to the specified  **CommandBarControl** control.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Execute**

 _expression_ Required. A variable that represents a **[CommandBarControl](commandbarcontrol-object-office.md)** object.


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


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

