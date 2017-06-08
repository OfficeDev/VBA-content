---
title: CommandBarButton.FaceId Property (Office)
keywords: vbaof11.chm6003
f1_keywords:
- vbaof11.chm6003
ms.prod: office
api_name:
- Office.CommandBarButton.FaceId
ms.assetid: c2151f20-b1c7-97eb-35ac-7a12c5ee3f28
ms.date: 06/08/2017
---


# CommandBarButton.FaceId Property (Office)

Gets or sets the Id number for the face of a  **CommandBarButton** control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **FaceId**

 _expression_ A variable that represents a **CommandBarButton** object.


## Remarks

The  **FaceId** property dictates the look, not the function, of a command bar button. The **Id** property of the **CommandBarControl** object determines the function of the button.

The value of the  **FaceId** property for a command bar button with a custom face is 0 (zero).


## Example

This example adds a command bar button to a custom command bar. Clicking this button is equivalent to clicking the  **Open** command on the **File** menu because the ID number is 23, yet the button has the same button face as the built-in **Charting** button.


```
Set newBar = CommandBars.Add(Name:="Custom2", _ 
     Position:=msoBarTop, Temporary:=True) 
newBar.Visible = True  
Set con = newBar.Controls.Add(Type:=msoControlButton, Id:=23) 
con.FaceId = 17
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

