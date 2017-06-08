---
title: MailingLabel.DefaultLaserTray Property (Word)
keywords: vbawd10.chm152502276
f1_keywords:
- vbawd10.chm152502276
ms.prod: word
api_name:
- Word.MailingLabel.DefaultLaserTray
ms.assetid: 0bc82fb0-abc3-7b46-c00b-8c009f2a6d91
ms.date: 06/08/2017
---


# MailingLabel.DefaultLaserTray Property (Word)

Returns or sets the default paper tray that contains sheets of mailing labels. Read/write  **WdPaperTray** .


## Syntax

 _expression_ . **DefaultLaserTray**

 _expression_ Required. A variable that represents a **[MailingLabel](mailinglabel-object-word.md)** object.


## Example

This example checks to determine whether the mailing label printer is set for feed labels manually, and then it displays a message on the status bar.


```vb
If Application.MailingLabel.DefaultLaserTray = _ 
 wdPrinterManualEnvelopeFeed Then 
 StatusBar = "Printer set for feeding labels manually" 
Else 
 StatusBar = "Check printer paper tray setting" 
End If
```

This example sets the mailing-label paper tray to the upper bin.




```vb
Application.MailingLabel.DefaultLaserTray = wdPrinterUpperBin
```


## See also


#### Concepts


[MailingLabel Object](mailinglabel-object-word.md)

