---
title: Application.RecordMacro Method (Excel)
keywords: vbaxl10.chm133195
f1_keywords:
- vbaxl10.chm133195
ms.prod: excel
api_name:
- Excel.Application.RecordMacro
ms.assetid: 8b6c9757-b589-04e6-5650-edfc4104e517
ms.date: 06/08/2017
---


# Application.RecordMacro Method (Excel)

Records code if the macro recorder is on.


## Syntax

 _expression_ . **RecordMacro**( **_BasicCode_** , **_XlmCode_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BasicCode_|Optional| **Variant**|A string that specifies the Visual Basic code that will be recorded if the macro recorder is recording into a Visual Basic module. The string will be recorded on one line. If the string contains a carriage return (ASCII character 10, or Chr$(10) in code), it will be recorded on more than one line.|
| _XlmCode_|Optional| **Variant**|This argument is ignored.|

## Remarks

The  **RecordMacro** method cannot record into the active module (the module in which the **RecordMacro** method exists).

If  _BasicCode_ is omitted and the application is recording into Visual Basic, Microsoft Excel will record a suitable `Application.Run` statement.

To prevent recording (for example, if the user cancels your dialog box), call this function with two empty strings.


## Example

This example records Visual Basic code.


```vb
Application.RecordMacro BasicCode:="Application.Run ""MySub"" "
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

