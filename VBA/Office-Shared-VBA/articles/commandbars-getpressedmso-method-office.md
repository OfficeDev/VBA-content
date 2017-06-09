---
title: CommandBars.GetPressedMso Method (Office)
keywords: vbaof11.chm2021
f1_keywords:
- vbaof11.chm2021
ms.prod: office
api_name:
- Office.CommandBars.GetPressedMso
ms.assetid: 97811bb6-cc5c-eccc-9149-76bdfa37541f
ms.date: 06/08/2017
---


# CommandBars.GetPressedMso Method (Office)

Returns a value indicating whether the toggleButton control identified by the  **idMso** parameter is pressed.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **GetPressedMso**( **_idMso_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return Value

Boolean


## Example

The following sample returns True when the  **Bold** button is pressed.


```
Application.CommandBars.GetPressedMso("Bold") 
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

