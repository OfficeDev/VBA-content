---
title: CommandBar.Visible Property (Office)
keywords: vbaof11.chm3020
f1_keywords:
- vbaof11.chm3020
ms.prod: office
api_name:
- Office.CommandBar.Visible
ms.assetid: c7057c83-ea8d-c167-a650-d784d5e6dd1f
ms.date: 06/08/2017
---


# CommandBar.Visible Property (Office)

Gets or sets the  **Visible** property of the command bar. **True** if the command bar is visible. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Visible**

 _expression_ A variable that represents a **CommandBar** object.


### Return Value

Boolean


## Remarks

The  **Visible** property for newly created custom command bars is **False** by default.

The  **Enabled** property for a command bar must be set to **True** before the **Visible** property is set to **True**.


## Example

This example steps through the collection of command bars to find the  **Forms** command bar. If the **Forms** command bar is found, the example makes it visible and protects its docking state.


```
foundFlag = False  
For Each cmdbar In CommandBars 
    If cmdbar.Name = "Forms" Then 
        cmdbar.Protection = msoBarNoChangeDock 
        cmdbar.Visible = True  
        foundFlag = True  
    End If 
Next 
If Not foundFlag Then 
    MsgBox "'Forms'command bar is not in the collection." 
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

