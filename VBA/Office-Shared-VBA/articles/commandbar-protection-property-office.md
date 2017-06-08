---
title: CommandBar.Protection Property (Office)
keywords: vbaof11.chm3015
f1_keywords:
- vbaof11.chm3015
ms.prod: office
api_name:
- Office.CommandBar.Protection
ms.assetid: 59f9e9d3-251c-93a6-fa49-75fa7c4f6659
ms.date: 06/08/2017
---


# CommandBar.Protection Property (Office)

Gets or sets a  **MsoBarProtection** constant representing the way a command bar is protected from user customization. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Protection**

 _expression_ A variable that represents a **CommandBar** object.


## Remarks

Using the constant  **msoBarNoCustomize** prevents users from accessing the **Add or Remove Buttons** menu (this menu enables users to customize a toolbar).


## Example

This example steps through the collection of command bars to find the command bar named "Forms." If this command bar is found, it's docking state is protected and it's made visible.


```
foundFlag =  False 
For i = 1 To CommandBars.Count 
    If CommandBars(i).Name = "Forms" Then 
            CommandBars(i).Protection = msoBarNoChangeDock 
            CommandBars(i).Visible = True  
            foundFlag = True  
    End If 
Next 
If Not foundFlag Then 
    MsgBox "'Forms' command bar is not in the collection." 
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

