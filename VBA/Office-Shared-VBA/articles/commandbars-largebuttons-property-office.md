---
title: CommandBars.LargeButtons Property (Office)
keywords: vbaof11.chm2009
f1_keywords:
- vbaof11.chm2009
ms.prod: office
api_name:
- Office.CommandBars.LargeButtons
ms.assetid: bcacab92-9779-5061-f68a-69722210e14e
ms.date: 06/08/2017
---


# CommandBars.LargeButtons Property (Office)

Is  **True** if the toolbar buttons displayed are larger than normal size. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **LargeButtons**

 _expression_ A variable that represents a **CommandBars** object.


## Example

This example switches the display size of toolbar buttons on all command bars.


```
Set allBars = CommandBars 
If allBars.LargeButtons Then 
    allBars.LargeButtons = False  
Else 
    allBars.LargeButtons = True  
End If
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

