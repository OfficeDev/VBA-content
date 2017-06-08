---
title: CommandBars.ReleaseFocus Method (Office)
keywords: vbaof11.chm2012
f1_keywords:
- vbaof11.chm2012
ms.prod: office
api_name:
- Office.CommandBars.ReleaseFocus
ms.assetid: 2ddca1e1-b8f4-a09c-120d-498b816747c4
ms.date: 06/08/2017
---


# CommandBars.ReleaseFocus Method (Office)

Releases the user interface focus from all command bars.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **ReleaseFocus**

 _expression_ A variable that represents a **CommandBars** object.


## Example

This example adds three blank buttons to the command bar named Custom and sets the focus to the center button. The example then waits five seconds before releasing the user interface focus from all command bars.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlButton 
    .Controls.Add Type:=msoControlButton 
    .Controls.Add Type:=msoControlButton 
    .Visible = True  
End With 
Set myControl = CommandBars("Custom").Controls(2) 
With myControl 
    .SetFocus 
End With 
PauseTime = 5   ' Set duration. 
    Start = Timer   ' Set start time. 
    Do While Timer  Start + PauseTime 
        DoEvents    ' Yield to other processes. 
    Loop 
    Finish = Timer 
CommandBars.ReleaseFocus
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

