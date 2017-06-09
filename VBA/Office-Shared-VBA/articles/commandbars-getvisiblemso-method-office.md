---
title: CommandBars.GetVisibleMso Method (Office)
keywords: vbaof11.chm2020
f1_keywords:
- vbaof11.chm2020
ms.prod: office
api_name:
- Office.CommandBars.GetVisibleMso
ms.assetid: ab916050-e1af-0752-9734-23d0fe27542f
ms.date: 06/08/2017
---


# CommandBars.GetVisibleMso Method (Office)

Returns True if the control identified by the  **idMso** parameter is visible.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **GetVisibleMso**( **_idMso_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return Value

Boolean


## Example

The following sample returns True if the  **Bold** button is visible.


```
Application.CommandBars.GetVisibleMso("Bold")
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

