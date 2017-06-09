---
title: WorksheetFunction.FilterXML Method (Excel)
keywords: vbaxl10.chm137465
f1_keywords:
- vbaxl10.chm137465
ms.prod: excel
ms.assetid: bcaa41a9-a122-ee87-29ca-cabb224358a1
ms.date: 06/08/2017
---


# WorksheetFunction.FilterXML Method (Excel)

Get specific data from the returned XML, typically from a  **WebService** function call.


## Syntax

 _expression_ . **FilterXML**_(Arg1,_ _Arg2)_

 _expression_ A variable that represents a object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required|STRING|Valid xml string.|
| _Arg2_|Required|STRING|XPath query string.|

### Remarks

The XPath parameter is limited to 1024 characters.

The  **FILTERXML** function returns results that are parsed via the user specified data locale.


### Return value

 **VARIANT**


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

