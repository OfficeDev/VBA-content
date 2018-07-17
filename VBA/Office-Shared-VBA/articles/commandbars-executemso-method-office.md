---
title: CommandBars.ExecuteMso Method (Office)
keywords: vbaof11.chm2018
f1_keywords:
- vbaof11.chm2018
ms.prod: office
api_name:
- Office.CommandBars.ExecuteMso
ms.assetid: 6f608475-7a79-48c7-abff-86d9ab07fe80
ms.date: 06/08/2017
---


# CommandBars.ExecuteMso Method (Office)

Executes the control identified by the  **idMso** parameter.


## Syntax

 _expression_. **ExecuteMso**( **_idMso_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Remarks

This method is useful in cases where there is no object model for a particular command. Works on controls that are built-in buttons, toggleButtons and splitButtons. On failure it returns E_InvalidArg for an invalid  **IdMso**, and E_Fail for controls that are not enabled or not visible.


## Example

The following sample executes the  **Copy** button.


```
Application.CommandBars.ExecuteMso("Copy")
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

