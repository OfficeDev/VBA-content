---
title: Document.ParseLine Method (Visio)
keywords: vis_sdr.chm10516425
f1_keywords:
- vis_sdr.chm10516425
ms.prod: visio
api_name:
- Visio.Document.ParseLine
ms.assetid: 46603de4-afa0-7903-f585-0a1aaa5c74c7
ms.date: 06/08/2017
---


# Document.ParseLine Method (Visio)

Parses a line of Microsoft Visual Basic code.


## Syntax

 _expression_ . **ParseLine**( **_Line_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Line_|Required| **String**|A string interpreted as Visual Basic code.|

### Return Value

Nothing


## Remarks

The  **ParseLine** method tells the Microsoft Visual Basic for Applications (VBA) project of the **Document** object to parse the string passed to it as an argument.

 The **ParseLine** method raises an exception if the string fails to parse.


## Example

You can use the following procedure to determine whether a string has been successfully parsed. If the parse succeeds, the procedure displays a message box that announces the success; if the parse fails, it displays a message box that announces the failure.


```vb
 
 
Public Sub ParseLine_Example() 
 
 Call LineParser("String to parse.") 
 
End Sub 
 
Public Sub LineParser(strString As String) 
 
 On Error Resume Next 
 ThisDocument.ParseLine strString 
 If Err = 0 Then 
 MsgBox "String parsed successfully" 
 Else 
 MsgBox "Parse not successful" 
 End If 
 
End Sub
```


