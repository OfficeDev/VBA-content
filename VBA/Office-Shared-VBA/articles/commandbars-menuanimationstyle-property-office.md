---
title: CommandBars.MenuAnimationStyle Property (Office)
keywords: vbaof11.chm2010
f1_keywords:
- vbaof11.chm2010
ms.prod: office
api_name:
- Office.CommandBars.MenuAnimationStyle
ms.assetid: bd79a55a-23f4-6056-649b-9dc384b597aa
ms.date: 06/08/2017
---


# CommandBars.MenuAnimationStyle Property (Office)

Gets or sets a  **MsoMenuAnimation** that represents the way a command bar is animated. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **MenuAnimationStyle**

 _expression_ A variable that represents a **CommandBars** object.


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

