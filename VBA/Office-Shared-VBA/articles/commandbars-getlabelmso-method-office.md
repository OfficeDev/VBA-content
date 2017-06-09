---
title: CommandBars.GetLabelMso Method (Office)
keywords: vbaof11.chm2022
f1_keywords:
- vbaof11.chm2022
ms.prod: office
api_name:
- Office.CommandBars.GetLabelMso
ms.assetid: 1ab6f700-e3c3-a89d-790f-10c27a6b495c
ms.date: 06/08/2017
---


# CommandBars.GetLabelMso Method (Office)

Returns the label of the control identified by the  **idMso** parameter as a String.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **GetLabelMso**( **_idMso_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return Value

String


## Example

The following sample returns the String "Paste".


```
Application.CommandBars.GetLabelMso("Paste")
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

