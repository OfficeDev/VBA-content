---
title: Application.RegisterXLL Method (Excel)
keywords: vbaxl10.chm133199
f1_keywords:
- vbaxl10.chm133199
ms.prod: excel
api_name:
- Excel.Application.RegisterXLL
ms.assetid: b0d97511-bb81-7c6a-7bbb-3f87c4364e95
ms.date: 06/08/2017
---


# Application.RegisterXLL Method (Excel)

Loads an XLL code resource and automatically registers the functions and commands contained in the resource.


## Syntax

 _expression_ . **RegisterXLL**( **_Filename_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|Specifies the name of the XLL to be loaded.|

### Return Value

Boolean


## Remarks

This method returns  **True** if the code resource is successfully loaded; otherwise, the method returns **False** .


## Example

This example loads an XLL file and registers the functions and commands in the file.


```vb
Application.RegisterXLL "XLMAPI.XLL"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

