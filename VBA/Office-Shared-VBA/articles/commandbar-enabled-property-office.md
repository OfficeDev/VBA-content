---
title: CommandBar.Enabled Property (Office)
keywords: vbaof11.chm3005
f1_keywords:
- vbaof11.chm3005
ms.prod: office
api_name:
- Office.CommandBar.Enabled
ms.assetid: 4a332d30-4aa9-1355-2d26-0d4f0529d488
ms.date: 06/08/2017
---


# CommandBar.Enabled Property (Office)

Gets or sets a  **Boolean** value that specifies whether the specified **CommandBar** is enabled. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Enabled**

 _expression_ A variable that represents a **[CommandBar](commandbar-object-office.md)** object.


## Remarks

For command bars, setting this property to  **True** causes the name of the command bar to appear in the list of available command bars.

For built-in controls, if you set the  **Enabled** property to **True**, the application determines its state, but setting it to **False** will force it to be disabled.


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

