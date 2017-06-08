---
title: CommandBars.ActionControl Property (Office)
keywords: vbaof11.chm2001
f1_keywords:
- vbaof11.chm2001
ms.prod: office
api_name:
- Office.CommandBars.ActionControl
ms.assetid: 70097691-a771-4f7d-020b-2a9d33e18fa0
ms.date: 06/08/2017
---


# CommandBars.ActionControl Property (Office)

Gets the  **CommandBarControl** object whose **OnAction** property is set to the running procedure. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **ActionControl**

 _expression_ A variable that represents a **CommandBars** object.


## Example

This example creates a command bar named "Custom", adds three buttons to it, and then uses the  **ActionControl** property and the **Tag** property to determine which command bar button was last clicked.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
Set buttonOne = myBar.Controls.Add(Type:=msoControlButton) 
With buttonOne 
    .FaceId = 133 
    .Tag = "RightArrow" 
    .OnAction = "whichButton" 
End With 
Set buttonTwo = myBar.Controls.Add(Type:=msoControlButton) 
With buttonTwo 
    .FaceId = 134 
    .Tag = "UpArrow" 
    .OnAction = "whichButton" 
End With 
Set buttonThree = myBar.Controls.Add(Type:=msoControlButton) 
With buttonThree 
    .FaceId = 135 
    .Tag = "DownArrow" 
    .OnAction = "whichButton" 
End With 
myBar.Visible = True
```

The following subroutine responds to the  **OnAction** method and determines which command bar button was last clicked.




```
Sub whichButton() 
Select Case CommandBars.ActionControl.Tag 
    Case "RightArrow" 
        MsgBox ("Right Arrow button clicked.") 
    Case "UpArrow" 
        MsgBox ("Up Arrow button clicked.") 
    Case "DownArrow" 
        MsgBox ("Down Arrow button clicked.") 
End Select 
End Sub
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

