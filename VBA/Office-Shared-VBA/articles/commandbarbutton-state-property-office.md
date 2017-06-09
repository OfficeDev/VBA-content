---
title: CommandBarButton.State Property (Office)
keywords: vbaof11.chm6006
f1_keywords:
- vbaof11.chm6006
ms.prod: office
api_name:
- Office.CommandBarButton.State
ms.assetid: 919ca064-507c-1db6-6b69-b586283ab67b
ms.date: 06/08/2017
---


# CommandBarButton.State Property (Office)

Gets or sets the appearance of a CommandBarButton control. Read/write.


## Syntax

 _expression_. **State**

 _expression_ Required. A variable that represents a **[CommandBarButton](commandbarbutton-object-office.md)** object.


## Remarks

The  **State** property of built-in command bar buttons is read-only. The value of the **Type** property is available as a value in the **[MsoButtonState](msobuttonstate-enumeration-office.md)** enumeration.


## Example

This example creates a command bar named "Custom" and adds two buttons to it. The example then sets the button on the left to  **msoButtonUp** and sets the button on the right to **msoButtonDown**.


```
 Dim myBar As Office.CommandBar 
 Dim imgSource As Office.CommandBarButton 
 Dim myControl1 As Office.CommandBarButton 
 Dim myControl2 As Office.CommandBarButton 
 ' Add new command bar. 
 Set myBar = CommandBars.Add(Name:="Custom", Position:=msoBarTop, Temporary:=True) 
 ' Add 2 buttons to new command bar. 
 With myBar 
 .Controls.Add Type:=msoControlButton 
 .Controls.Add Type:=msoControlButton 
 .Visible = True 
 End With 
 ' Paste Bold button face and set State of first button. 
 Set myControl1 = myBar.Controls(1) 
 Set imgSource = CommandBars.FindControl(msoControlButton, 113) 
 imgSource.CopyFace 
 With myControl1 
 .PasteFace 
 .State = msoButtonUp 
 End With 
 ' Paste italic button face and set State of second button. 
 Set myControl2 = myBar.Controls(2) 
 Set imgSource = CommandBars.FindControl(msoControlButton, 114) 
 imgSource.CopyFace 
 With myControl2 
 .PasteFace 
 .State = msoButtonDown 
 End With 

```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

