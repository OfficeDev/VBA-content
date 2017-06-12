---
title: CommandBarControl.OnAction Property (Office)
ms.prod: office
api_name:
- Office.CommandBarControl.OnAction
ms.assetid: 05e40fcb-ff67-049f-6386-a9ef20b48c87
ms.date: 06/08/2017
---


# CommandBarControl.OnAction Property (Office)

Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a  **CommandBarControl**. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **OnAction**

 _expression_ A variable that represents a **CommandBarControl** object.


### Return Value

String


## Remarks

The container application determines whether the value is a valid macro name.


## Example

This example adds a command bar control to the command bar named "Custom". The procedure named  **MySub** will run each time the control is clicked.


```
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls _ 
    .Add(Type:=msocontrolButton) 
With myControl 
    .FaceId = 2 
    .OnAction = "MySub" 
End With 
myBar.Visible = True
```

This example adds a command bar control to the command bar named "Custom". The COM add-in named "FinanceAddIn" will run each time the control is clicked.




```
Set myBar = CommandBars("Custom") 
Set myControl = myBar.Controls _ 
    .Add(Type:=msocontrolButton) 
With myControl 
    .FaceId = 2 
    .OnAction = "!<FinanceAddIn>" 
End With 
myBar.Visible = True
```


## See also


#### Concepts


[CommandBarControl Object](commandbarcontrol-object-office.md)
#### Other resources


[CommandBarControl Object Members](commandbarcontrol-members-office.md)

