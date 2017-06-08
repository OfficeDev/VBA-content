---
title: Application.DDEPoke Method (Word)
keywords: vbawd10.chm158335288
f1_keywords:
- vbawd10.chm158335288
ms.prod: word
api_name:
- Word.Application.DDEPoke
ms.assetid: b782fc34-551f-288f-e087-5429f7ee7814
ms.date: 06/08/2017
---


# Application.DDEPoke Method (Word)

Uses an open dynamic data exchange (DDE) channel to send data to an application.


## Syntax

 _expression_ . **DDEPoke**( **_Channel_** , **_Item_** , **_Data_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **DDEInitiate** method.|
| _Item_|Required| **String**|The item within a DDE topic to which the specified data is to be sent.|
| _Data_|Required| **String**|The data to be sent to the receiving application (the DDE server).|

## Remarks


 **Security Note**  



If the  **DDEPoke** method isn't successful, an error occurs.


## Example

This example opens the Microsoft Excel workbook Sales.xls and inserts "1996 Sales" into cell R1C1.


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[OPEN(" &; Chr(34) _ 
 &; "C:\Sales.xls" &; Chr(34) &; ")] 
DDETerminate Channel:=lngChannel 
lngChannel = DDEInitiate(App:="Excel", Topic:="Sales.xls") 
DDEPoke Channel:=lngChannel, Item:="R1C1", Data:="1996 Sales" 
DDETerminate Channel:=lngChannel
```


## See also


#### Concepts


[Application Object](application-object-word.md)

