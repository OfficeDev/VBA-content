---
title: CommandBarButton.PasteFace Method (Office)
keywords: vbaof11.chm6004
f1_keywords:
- vbaof11.chm6004
ms.prod: office
api_name:
- Office.CommandBarButton.PasteFace
ms.assetid: 1c4179c4-b6b5-527f-5027-25ced8ee907d
ms.date: 06/08/2017
---


# CommandBarButton.PasteFace Method (Office)

Pastes the contents of the Clipboard onto a  **CommandBarButton**.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **PasteFace**

 _expression_ A variable that represents a **CommandBarButton** object.


## Example

This example finds the built-in  **FileOpen** button and pastes the face from the **Spelling** and **Grammar** button onto it from the Clipboard.


```
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=2) 
myControl.CopyFace 
Set myControl = CommandBars.FindControl(Type:=msoControlButton, Id:=23) 
myControl.PasteFace
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

