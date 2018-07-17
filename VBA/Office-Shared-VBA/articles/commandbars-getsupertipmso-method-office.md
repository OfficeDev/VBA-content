---
title: CommandBars.GetSupertipMso Method (Office)
keywords: vbaof11.chm2024
f1_keywords:
- vbaof11.chm2024
ms.prod: office
api_name:
- Office.CommandBars.GetSupertipMso
ms.assetid: e116402f-bbb7-8cd3-6305-7daf85feb514
ms.date: 06/08/2017
---


# CommandBars.GetSupertipMso Method (Office)

Returns the supertip of the control identified by the  **idMso** parameter as a String.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **GetSupertipMso**( **_idMso_** )

 _expression_ An expression that returns a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return Value

String


## Example

The following sample returns the String "Cut the selection from the document and put it on the Clipboard."


```
Application.CommandBars.GetSupertipMso("Cut")
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

