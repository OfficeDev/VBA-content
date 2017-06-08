---
title: CommandBars.Item Property (Office)
keywords: vbaof11.chm2008
f1_keywords:
- vbaof11.chm2008
ms.prod: office
api_name:
- Office.CommandBars.Item
ms.assetid: bca38d83-67cb-2cba-ddfa-918a5b2ff508
ms.date: 06/08/2017
---


# CommandBars.Item Property (Office)

Gets a  **CommandBar** object from the **CommandBars** collection. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[CommandBars](commandbars-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the object to be returned.|

## Example

Item is the default member of the object or collection. The following two statements both assign a CommandBar object to cmdBar.


```
Set cmdBar = CommandBars.Item("Standard") 
Set cmdBar = CommandBars("Standard")
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

