---
title: CommandBar.Context Property (Office)
keywords: vbaof11.chm3002
f1_keywords:
- vbaof11.chm3002
ms.prod: office
api_name:
- Office.CommandBar.Context
ms.assetid: e7b8a7e5-0799-84e8-c7e3-5f713971099d
ms.date: 06/08/2017
---


# CommandBar.Context Property (Office)

Gets or sets a string that determines where a command bar will be saved. The string is defined and interpreted by the application. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Context**

 _expression_ A variable that represents a **CommandBar** object.


## Remarks

You can set the  **Context** property only for custom command bars. This property will fail if the application doesn't recognize the context string, or if the application doesn't support changing context strings programmatically.


## Example

This example displays a message box containing the context string for the command bar named "Custom". This example works in Microsoft Word and other applications that support the  **Context** property.


```
Set myBar = CommandBars _ 
    .Add(Name:="Custom", Position:=msoBarTop, _ 
    Temporary:=True) 
With myBar 
    .Controls.Add Type:=msoControlButton, ID:=2 
    .Visible = True  
End With 
MsgBox (myBar.Context) 

```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

