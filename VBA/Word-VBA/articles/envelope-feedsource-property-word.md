---
title: Envelope.FeedSource Property (Word)
keywords: vbawd10.chm152567820
f1_keywords:
- vbawd10.chm152567820
ms.prod: word
api_name:
- Word.Envelope.FeedSource
ms.assetid: c6794e83-8136-7e50-fa82-819d4d6d6f8b
ms.date: 06/08/2017
---


# Envelope.FeedSource Property (Word)

Returns or sets the paper tray for the envelope. Read/write  **WdPaperTray** .


## Syntax

 _expression_ . **FeedSource**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Example

This example asks the user whether envelopes are fed into the printer manually. If the answer is yes, the example sets the paper tray to manual envelope feed.


```vb
Sub exFeedSource() 
 
 Dim intResponse As Integer 
 
 intResponse = _ 
 MsgBox("Are the envelopes manually fed?", vbYesNo) 
 If intResponse = vbYes then 
 On Error GoTo errhandler 
 ActiveDocument.Envelope.FeedSource = _ 
 wdPrinterManualEnvelopeFeed 
 End If 
 
 Exit Sub 
 
errhandler: 
 If Err = 5852 Then MsgBox _ 
 "Envelope not part of the active document" 
 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

