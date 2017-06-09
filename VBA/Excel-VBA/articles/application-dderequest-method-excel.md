---
title: Application.DDERequest Method (Excel)
keywords: vbaxl10.chm183093
f1_keywords:
- vbaxl10.chm183093
ms.prod: excel
api_name:
- Excel.Application.DDERequest
ms.assetid: 822ef77e-5f11-aced-f770-05175ce128c7
ms.date: 06/08/2017
---


# Application.DDERequest Method (Excel)

Requests information from the specified application. This method always returns an array.


## Syntax

 _expression_ . **DDERequest**( **_Channel_** , **_Item_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **[DDEInitiate](application-ddeinitiate-method-excel.md)** method.|
| _Item_|Required| **String**|The item to be requested.|

### Return Value

Variant


## Example

This example opens a channel to the System topic in Word and then uses the Topics item to return a list of all open documents. The list is returned in column A on Sheet1.


```vb
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="System") 
returnList = Application.DDERequest(channelNumber, "Topics") 
For i = LBound(returnList) To UBound(returnList) 
 Worksheets("Sheet1").Cells(i, 1).Formula = returnList(i) 
Next i 
Application.DDETerminate channelNumber
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

