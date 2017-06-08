---
title: CommandBars.Add Method (Office)
keywords: vbaof11.chm2003
f1_keywords:
- vbaof11.chm2003
ms.prod: office
api_name:
- Office.CommandBars.Add
ms.assetid: 544cfa94-924a-90ca-d716-c7b2f9e8732f
ms.date: 06/08/2017
---


# CommandBars.Add Method (Office)

Creates a new command bar and adds it to the collection of command bars.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Add**( **_Name_**, **_Position_**, **_MenuBar_**, **_Temporary_** )

 _expression_ Required. A variable that represents a **[CommandBars](commandbars-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**Variant**|The name of the new command bar. If this argument is omitted, a default name is assigned to the command bar (such as Custom 1).|
| _Position_|Optional|**Variant**|The position or type of the new command bar. Can be one of the  **MsoBarPosition** constants.|
| _MenuBar_|Optional|**Variant**|**True** to replace the active menu bar with the new command bar. The default value is **False**.|
| _Temporary_|Optional|**Variant**|**True** to make the new command bar temporary. Command bars are deleted when the container application is closed. The default value is **False**.|

## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

