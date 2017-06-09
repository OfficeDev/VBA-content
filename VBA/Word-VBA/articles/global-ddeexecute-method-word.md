---
title: Global.DDEExecute Method (Word)
keywords: vbawd10.chm163119414
f1_keywords:
- vbawd10.chm163119414
ms.prod: word
api_name:
- Word.Global.DDEExecute
ms.assetid: f1ae9960-59b9-e637-c507-606c13e36c92
ms.date: 06/08/2017
---


# Global.DDEExecute Method (Word)

Sends a command or series of commands to an application through the specified dynamic data exchange (DDE) channel.


## Syntax

 _expression_ . **DDEExecute**( **_Channel_** , **_Command_** )

 _expression_ A variable that represents a **[Global](global-object-word.md)** object. Optional.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **DDEInitiate** method.|
| _Command_|Required| **String**|A command or series of commands recognized by the receiving application (the DDE server). If the receiving application cannot perform the specified command, an error occurs.|

## Remarks


 **Security Note**  




## Example

This example creates a new worksheet in Microsoft Excel. The XLM macro instruction to create a new worksheet is New(1).


```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[New(1)]" 
DDETerminate Channel:=lngChannel
```

This example runs the Microsoft Excel macro named "Macro1" in Personal.xls.




```vb
Dim lngChannel As Long 
 
lngChannel = DDEInitiate(App:="Excel", Topic:="System") 
DDEExecute Channel:=lngChannel, Command:="[Run(" &; Chr(34) &; _ 
 "Personal.xls!Macro1" &; Chr(34) &; ")]" 
DDETerminate Channel:=lngChannel
```


## See also


#### Concepts


[Global Object](global-object-word.md)

