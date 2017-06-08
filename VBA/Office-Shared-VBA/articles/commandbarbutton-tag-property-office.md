---
title: CommandBarButton.Tag Property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Tag
ms.assetid: c73a12a8-8b20-1e32-ad98-ae0bb3b1daed
ms.date: 06/08/2017
---


# CommandBarButton.Tag Property (Office)

Gets or sets information about the  **CommandBarButton** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Tag**

 _expression_ A variable that represents a **CommandBarButton** object.


### Return Value

String


## Example

To avoid duplicate calls of the same class when tiggered with events, define the  **Tag** property unique to the events. The following example demonstrates this concept with two modules.


```
Public WithEvents oBtn As CommandBarButton 
 
Private Sub oBtn_click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean) 
    MsgBox "Clicked " &amp; ctrl.Caption 
 
End Sub 
 
Dim oBtns As New Collection 
      
Sub Use_Tag() 
     
    Dim oEvt As CBtnEvent 
    Set oBtns = Nothing 
 
    For i = 1 To 5 
        Set oEvt = New CBtnEvent 
        Set oEvt.oBtn = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlButton) 
        With oEvt.oBtn 
            .Caption = "Btn" &amp; i 
            .Style = msoButtonCaption 
            .Tag = "Hello" &amp; i 
        End With 
        oBtns.Add oEvt 
    Next 
      
End Sub
```

This example sets the tag for the button on the custom command bar to "Spelling Button" and displays the tag in a message box.




```
CommandBars("Custom").Controls(1).Tag = "Spelling Button" 
MsgBox (CommandBars("Custom").Controls(1).Tag)
```


## See also


#### Concepts


[CommandBarButton Object](commandbarbutton-object-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

