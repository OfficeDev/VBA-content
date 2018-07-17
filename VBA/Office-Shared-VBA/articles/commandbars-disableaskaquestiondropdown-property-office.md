---
title: CommandBars.DisableAskAQuestionDropdown Property (Office)
keywords: vbaof11.chm2017
f1_keywords:
- vbaof11.chm2017
ms.prod: office
api_name:
- Office.CommandBars.DisableAskAQuestionDropdown
ms.assetid: a0954aa4-256c-4a14-6bab-959a00e9367d
ms.date: 06/08/2017
---


# CommandBars.DisableAskAQuestionDropdown Property (Office)

Is  **True** if the **Answer Wizard** dropdown menu is enabled. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **DisableAskAQuestionDropdown**

 _expression_ A variable that represents a **CommandBars** object.


## Example

The following example switches the  **DisableAskAQuestionDropdown** property on or off.


```
Sub ToggleQuestionDropdown() 
    With Application.CommandBars 
        If .DisableAskAQuestionDropdown =  True Then 
            .DisableAskAQuestionDropdown = False  
        Else 
            .DisableAskAQuestionDropdown = True  
        End If 
    End With 
End Sub
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

