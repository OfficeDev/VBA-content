---
title: Speech.SpeakCellOnEnter Property (Excel)
keywords: vbaxl10.chm718075
f1_keywords:
- vbaxl10.chm718075
ms.prod: excel
api_name:
- Excel.Speech.SpeakCellOnEnter
ms.assetid: a176820a-85ef-338c-b507-9ffb9d744631
ms.date: 06/08/2017
---


# Speech.SpeakCellOnEnter Property (Excel)

Microsoft Excel supports a mode where the active cell will be spoken when the ENTER key is pressed or when the active cell is finished being edited. Setting the  **SpeakCellOnEnter** property to **True** will turn this mode on. **False** turns this mode off. Read/write **Boolean** .


## Syntax

 _expression_ . **SpeakCellOnEnter**

 _expression_ A variable that represents a **Speech** object.


## Example

This example determines if the active cell will be spoken when the ENTER key is pressed or the active cell is done being edited, and notifies the user.


```vb
Sub SpeechCheck() 
 
 ' Determine mode setting and notify user. 
 If Application.Speech.SpeakCellOnEnter = True Then 
 MsgBox "The Speak On Enter mode is turned on. " &; _ 
 "The active cell will be spoken when the ENTER "&; _ 
 "key is pressed or it is done being edited." 
 Else 
 MsgBox "The Speaker On Enter mode is turned off." 
 End If 
 
End Sub
```


## See also


#### Concepts


[Speech Object](speech-object-excel.md)

