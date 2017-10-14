---
title: CommandBarButton.CopyFace Method (Office)
keywords: vbaof11.chm6002
f1_keywords:
- vbaof11.chm6002
ms.prod: office
api_name:
- Office.CommandBarButton.CopyFace
ms.assetid: 09f09dbd-b70f-8b7d-1af7-7e43bffe3030
ms.date: 06/08/2017
---


# CommandBarButton.CopyFace Method (Office)

Copies the face of a command bar button control to the Clipboard.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **CopyFace**

 _expression_ A variable that represents a **CommandBarButton** object.


## Remarks

Use the  **PasteFace** method to paste the contents of the Clipboard onto a button face.


## Example

This example finds the built-in  **Open** button, copies the button face to the Clipboard, and then pastes the face onto the **Spelling** and **Grammar** button.


```
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=23) 
myControl.CopyFace 
Set myControl = CommandBars.FindControl(Type:=msoControlButton, ID:=2) 
myControl.PasteFace
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

