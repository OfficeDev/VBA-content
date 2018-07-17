---
title: Global.DDETerminate Method (Word)
keywords: vbawd10.chm163119418
f1_keywords:
- vbawd10.chm163119418
ms.prod: word
api_name:
- Word.Global.DDETerminate
ms.assetid: 2502d0a7-c90b-1169-7b7b-a5d2b26445a6
ms.date: 06/08/2017
---


# Global.DDETerminate Method (Word)

Closes the specified dynamic data exchange (DDE) channel to another application.


## Syntax

 _expression_ . **DDETerminate**( **_Channel_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **DDEInitiate** method.|

## Remarks


 **Security Note**  




## Example

This example creates a new workbook in Microsoft Excel and then terminates the DDE conversation.


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[New(1)]" 
DDETerminate Channel:=lngChannel
```


## See also


#### Concepts


[Global Object](global-object-word.md)

