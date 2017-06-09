---
title: CommandBarPopup.Caption Property (Office)
ms.prod: office
api_name:
- Office.CommandBarPopup.Caption
ms.assetid: fc9221e6-cfb0-9f2a-290b-73a434569e65
ms.date: 06/08/2017
---


# CommandBarPopup.Caption Property (Office)

Gets or sets the caption text for a command bar control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Caption**

 _expression_ A variable that represents a **CommandBarPopup** object.


### Return Value

String


## Example

This example adds a command bar control with a spelling checker button face to a custom command bar, and then it sets the caption to "Spelling checker."


```
Set myBar = CommandBars.Add(Name:="Custom", _ 
Position:=msoBarTop, Temporary:=True) 
myBar.Visible = True  
Set myControl = myBar.Controls _ 
.Add(Type:=msoControlButton, Id:=2) 
With myControl 
    .DescriptionText = "Starts the spelling checker" 
    .Caption = "Spelling checker" 
End With
```


 **Note**  


 **Note**  A control's caption is also displayed as its default ScreenTip.


## See also


#### Concepts


[CommandBarPopup Object](commandbarpopup-object-office.md)
#### Other resources


[CommandBarPopup Object Members](commandbarpopup-members-office.md)

