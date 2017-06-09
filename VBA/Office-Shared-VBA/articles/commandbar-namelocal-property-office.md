---
title: CommandBar.NameLocal Property (Office)
keywords: vbaof11.chm3011
f1_keywords:
- vbaof11.chm3011
ms.prod: office
api_name:
- Office.CommandBar.NameLocal
ms.assetid: 3afad045-aaf8-8775-574e-faaccde7d270
ms.date: 06/08/2017
---


# CommandBar.NameLocal Property (Office)

Gets the name of a built-in command bar as it's displayed in the language version of the container application, or returns or sets the name of a custom command bar. Read/write.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **NameLocal**

 _expression_ A variable that represents a **CommandBar** object.


## Remarks


 **Note**  If you attempt to set this property for a built-in command bar, an error occurs.

The local name of a built-in command bar is displayed in the title bar (when the command bar isn't docked) and in the list of available command bars, wherever that list is displayed in the container application.

If you change the value of the  **LocalName** property for a custom command bar, the value of **Name** changes as well, and vice versa.


## Example

This example displays the name and localized name of the first command bar in the container application.


```
With CommandBars(1) 
    MsgBox "The name of the command bar is " &amp; .Name 
    MsgBox "The localized name of the command bar is " &amp; .NameLocal 
End With
```


## See also


#### Concepts


[CommandBar Object](commandbar-object-office.md)
#### Other resources


[CommandBar Object Members](commandbar-members-office.md)

