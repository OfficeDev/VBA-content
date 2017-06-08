---
title: Global.DDETerminateAll Method (Word)
keywords: vbawd10.chm163119419
f1_keywords:
- vbawd10.chm163119419
ms.prod: word
api_name:
- Word.Global.DDETerminateAll
ms.assetid: 67688b3e-06df-a6b0-10f6-ecbd29ed386a
ms.date: 06/08/2017
---


# Global.DDETerminateAll Method (Word)

Closes all dynamic data exchange (DDE) channels opened by Microsoft Word. .


## Syntax

 _expression_ . **DDETerminateAll**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


## Remarks

This method does not close channels opened to Word by client applications. Using this method is the same as using the  **DDETerminate** method for each open channel.


 **Security Note**  



If you interrupt a macro that opens a DDE channel, you may inadvertently leave a channel open. Open channels are not closed automatically when a macro ends, and each open channel uses system resources. For this reason, it is a good idea to use this method when you are debugging a macro that opens one or more DDE channels.


## Example

This example opens the Microsoft Office Excel workbook Book1.xls, inserts text into cell R2C3, saves the workbook, and then terminates all DDE channels.


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[OPEN(" &; Chr(34) &; _ 
 "C:\Documents\Book1.xls" &; Chr(34) &; ")]" 
DDETerminate Channel:=lngChannel 
lngChannel = DDEInitiate(App:="Excel", Topic:="Book1.xls") 
DDEPoke Channel:=lngChannel, Item:="R2C3", Data:="Hello World" 
DDEExecute Channel:=lngChannel, Command:="[Save]" 
DDETerminateAll
```


## See also


#### Concepts


[Global Object](global-object-word.md)

