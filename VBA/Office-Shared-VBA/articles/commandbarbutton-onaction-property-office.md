---
title: CommandBarButton.OnAction Property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.OnAction
ms.assetid: c0a4148c-330a-6bd9-dd14-7ade8fc833fe
ms.date: 06/08/2017
---


# CommandBarButton.OnAction Property (Office)

Gets or sets the name of a Visual Basic procedure that will run when the user clicks or changes the value of a  **CommandBarButton** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **OnAction**

 _expression_ A variable that represents a **CommandBarButton** object.


### Return Value

String


## Remarks

The container application determines whether the value is a valid macro name.


## Example

This example adds a command bar control to the command bar named "Custom". The COM add in named "FinanceAddIn" will run each time the control is clicked.


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


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

