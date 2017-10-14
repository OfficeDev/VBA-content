---
title: Global.DDEInitiate Method (Word)
keywords: vbawd10.chm163119415
f1_keywords:
- vbawd10.chm163119415
ms.prod: word
api_name:
- Word.Global.DDEInitiate
ms.assetid: 4b27c9dc-6d81-50e7-968b-f583cd1f23b9
ms.date: 06/08/2017
---


# Global.DDEInitiate Method (Word)

Opens a dynamic data exchange (DDE) channel to another application, and returns the channel number.


## Syntax

 _expression_ . **DDEInitiate**( **_App_** , **_Topic_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _App_|Required| **String**|The name of the application.|
| _Topic_|Required| **String**|The name of a DDE topic?for example, the name of an open document?recognized by the application to which you are opening a channel.|

## Remarks


 **Security Note**  



If it is successful, the  **DDEInitiate** method returns the number of the open channel. All subsequent DDE functions use this number to specify the channel.


## Example

This example initiates a DDE conversation with the System topic and opens the Microsoft Office Excel workbook Sales.xls. The example terminates the DDE channel, initiates a channel to Sales.xls, and then inserts text into cell R1C1.


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


[Global Object](global-object-word.md)

