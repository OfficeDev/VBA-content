---
title: CommandBar.Position Property (Office)
keywords: vbaof11.chm3013
f1_keywords:
- vbaof11.chm3013
ms.prod: office
api_name:
- Office.CommandBar.Position
ms.assetid: b1e80bc0-1586-523b-a9ec-70c76fa54252
ms.date: 06/08/2017
---


# CommandBar.Position Property (Office)

Gets or sets a  **MsoBarPosition** constant representing the position of a command bar. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Position**

 _expression_ A variable that represents a **CommandBar** object.


## Example

This example steps through the collection of command bars, docking the custom command bars at the bottom of the application window and docking the built-in command bars at the top of the window.


```
For Each bar In CommandBars 
    If bar.Visible = True Then 
        If bar.BuiltIn Then 
            bar.Position = msoBarTop 
         Else 
            bar.Position = msoBarBottom 
        End If 
    End If 
Next
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

