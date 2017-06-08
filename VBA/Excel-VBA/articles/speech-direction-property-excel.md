---
title: Speech.Direction Property (Excel)
keywords: vbaxl10.chm718074
f1_keywords:
- vbaxl10.chm718074
ms.prod: excel
api_name:
- Excel.Speech.Direction
ms.assetid: 8cbedcb3-2d91-b9c1-c1ae-6f06cd8d442b
ms.date: 06/08/2017
---


# Speech.Direction Property (Excel)

Returns or sets the order in which the cells will be spoken. The value of the  **Direction** property is an **[XlSpeakDirection](xlspeakdirection-enumeration-excel.md)** constant. Read/write.


## Syntax

 _expression_ . **Direction**

 _expression_ A variable that represents a **Speech** object.


## Remarks





| **XlSpeakDirection** can be one of these **XlSpeakDirection** constants.|
| **xlSpeakByColumns**|
| **xlSpeakByRows**|

## Example

In this example, Microsoft Excel determines the speech direction and notifies the user.


```vb
Sub CheckSpeechDirection() 
 
 ' Notify user of speech direction. 
 If Application.Speech.Direction = xlSpeakByColumns Then 
 MsgBox "The speech direction is set to speak by columns." 
 Else 
 MsgBox "The speech direction is set to speak by rows." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Databar Object](databar-object-excel.md)
[Speech Object](speech-object-excel.md)

