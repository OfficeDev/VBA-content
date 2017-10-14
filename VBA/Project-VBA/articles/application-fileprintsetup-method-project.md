---
title: Application.FilePrintSetup Method (Project)
keywords: vbapj.chm113
f1_keywords:
- vbapj.chm113
ms.prod: project-server
api_name:
- Project.Application.FilePrintSetup
ms.assetid: 87c49847-3b00-28d7-f45b-3205947a6627
ms.date: 06/08/2017
---


# Application.FilePrintSetup Method (Project)

Specifies the active printer.


## Syntax

 _expression_. **FilePrintSetup**( ** _Printer_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Printer_|Optional|**String**|The full name or port name of the active printer.|

### Return Value

 **Boolean**


## Example

The following example sets the active printer to the printer on the LPT1 port.


```vb
Sub SetActivePrinterToLPT1() 
 FilePrintSetup "LPT1:" 
End Sub
```


