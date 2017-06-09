---
title: CommandBarButton Object (Office)
keywords: vbaof11.chm244000
f1_keywords:
- vbaof11.chm244000
ms.prod: office
api_name:
- Office.CommandBarButton
ms.assetid: e6d8209d-2c87-f1b5-bc3f-d4e5e5d3ab73
ms.date: 06/08/2017
---


# CommandBarButton Object (Office)

Represents a button control on a command bar.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Example

Use  **Controls(index)**, where _index_ is the index number of the control, to return a **CommandBarButton** object. Note that the **Type** property of the control must be **msoControlButton**. Assuming that the second control on the command bar named "Custom" is a button, the following example changes the style of that button.


```
Set c = CommandBars("Custom").Controls(2) 
With c 
If .Type = msoControlButton Then 
    If .Style = msoButtonIcon Then 
        .Style = msoButtonIconAndCaption 
    Else 
        .Style = msoButtonIcon 
    End If 
End If 
End With
```


 **Note**  


 **Note**  You can also use the  **FindControl** method to return a **CommandBarButton** object.


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[CommandBarButton Object Members](commandbarbutton-members-office.md)

