---
title: Global.DDERequest Method (Word)
keywords: vbawd10.chm163119417
f1_keywords:
- vbawd10.chm163119417
ms.prod: word
api_name:
- Word.Global.DDERequest
ms.assetid: be540a7b-9a38-633a-cf48-2a15a3159a51
ms.date: 06/08/2017
---


# Global.DDERequest Method (Word)

Uses an open dynamic data exchange (DDE) channel to request information from the receiving application, and returns the information as a string.


## Syntax

 _expression_ . **DDERequest**( **_Channel_** , **_Item_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **DDEInitiate** method.|
| _Item_|Required| **String**|The item to be requested.|

## Remarks


 **Security Note**  



When you request information from the topic in the server application, you must specify the item in that topic whose contents you are requesting. In Microsoft Excel, for example, cells are valid items, and you refer to them by using either the "R1C1" format or named references.

Microsoft Excel and other applications that support DDE recognize a topic named "System." Three standard items in the System topic are described in the following table. Note that you can get a list of the other items in the System topic by using the SysItems item.



|**Item in System topic**|**Effect**|
|:-----|:-----|
|SysItems|Returns a list of all the items in the System topic.|
|Topics|Returns a list of all the available topics.|
|Formats|Returns a list of all the Clipboard formats supported by Word.|

## Example

This example opens the Microsoft Excel workbook Book1.xls and retrieves the contents of cell R1C1.


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[OPEN(" &; Chr(34) _ 
 &; "C:\Documents\Book1.xls" &; Chr(34) &; ")]" 
DDETerminate Channel:=lngChannel 
lngChannel = DDEInitiate(App:="Excel", Topic:="Book1.xls") 
MsgBox DDERequest(Channel:=lngChannel, Item:="R1C1") 
DDETerminateAll
```

This example opens a channel to the System topic in Microsoft Excel and then uses the Topics item to return a list of available topics. The example inserts the topic list, which includes all open workbooks, after the selection.




```vb
Dim lngChannel As Long 
Dim strTopicList As String 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
strTopicList = DDERequest(Channel:=lngChannel, Item:="Topics") 
Selection.InsertAfter strTopicList 
DDETerminate Channel:=lngChannel
```


## See also


#### Concepts


[Global Object](global-object-word.md)

