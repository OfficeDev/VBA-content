---
title: CommandBar.BuiltIn Property (Office)
keywords: vbaof11.chm3001
f1_keywords:
- vbaof11.chm3001
ms.prod: office
api_name:
- Office.CommandBar.BuiltIn
ms.assetid: f7e4c581-2019-9fca-5e9e-15db4d656269
ms.date: 06/08/2017
---


# CommandBar.BuiltIn Property (Office)

Gets  **True** if the specified command bar is a built-in command bar of the container application. Returns **False** if it is a custom command bar. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **BuiltIn**

 _expression_ A variable that represents a **CommandBar** object.


### Return Value

Boolean


## Example

This example deletes all custom command bars that aren't visible.


```
foundFlag = False  
deletedBars = 0 
For Each bar In CommandBars 
    If (bar.BuiltIn = False) And (bar.Visible = False) Then 
        bar.Delete 
        foundFlag = True  
        deletedBars = deletedBars + 1 
    End If 
Next 
If Not foundFlag Then 
    MsgBox "No command bars have been deleted." 
Else 
    MsgBox deletedBars &amp; " custom command bar(s) deleted." 
End If
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

