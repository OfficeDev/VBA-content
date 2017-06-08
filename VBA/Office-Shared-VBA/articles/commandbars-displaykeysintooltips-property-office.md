---
title: CommandBars.DisplayKeysInTooltips Property (Office)
keywords: vbaof11.chm2006
f1_keywords:
- vbaof11.chm2006
ms.prod: office
api_name:
- Office.CommandBars.DisplayKeysInTooltips
ms.assetid: de132c5f-bc9f-c335-28ff-b9459c912b2c
ms.date: 06/08/2017
---


# CommandBars.DisplayKeysInTooltips Property (Office)

Is  **True** if shortcut keys are displayed in the **ToolTips** for each command bar control. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **DisplayKeysInTooltips**

 _expression_ A variable that represents a **CommandBars** object.


## Remarks

To display shortcut keys in  **ToolTips**, you must also set the  **DisplayTooltips** property to **True**.


## Example

This example sets options for all command bars in Microsoft Office.


```
With CommandBars 
    .LargeButtons = True  
    .DisplayTooltips = True  
    .DisplayKeysInTooltips = True  
    .MenuAnimationStyle = msoMenuAnimationUnfold 
End With
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

