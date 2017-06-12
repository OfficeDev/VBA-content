---
title: CommandBars.DisableCustomize Property (Office)
keywords: vbaof11.chm2016
f1_keywords:
- vbaof11.chm2016
ms.prod: office
api_name:
- Office.CommandBars.DisableCustomize
ms.assetid: cbebdaa7-2e8d-af73-fd18-03b3b11f98ac
ms.date: 06/08/2017
---


# CommandBars.DisableCustomize Property (Office)

Is  **True** if toolbar customization is disabled. Read/write.


## Syntax

 _expression_. **DisableCustomize**

 _expression_ A variable that represents a **CommandBars** object.


## Example

The following example switches the  **DisableCustomize** property on or off.


```
Sub ToggleCustomize() 
 With Application.CommandBars 
 If .DisableCustomize = True Then 
 .DisableCustomize = False 
 Else 
 .DisableCustomize = True 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

