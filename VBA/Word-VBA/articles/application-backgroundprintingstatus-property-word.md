---
title: Application.BackgroundPrintingStatus Property (Word)
keywords: vbawd10.chm158335062
f1_keywords:
- vbawd10.chm158335062
ms.prod: word
api_name:
- Word.Application.BackgroundPrintingStatus
ms.assetid: 74fabdd0-55d8-63c6-4608-36af8138b3c1
ms.date: 06/08/2017
---


# Application.BackgroundPrintingStatus Property (Word)

Returns the number of print jobs in the background printing queue. Read-only  **Long** .


## Syntax

 _expression_ . **BackgroundPrintingStatus**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example returns the number of Word print jobs currently queued up for background printing.


```vb
Dim lngStatus As Long 
 
If Options.PrintBackground = True Then 
 lngStatus = Application.BackgroundPrintingStatus 
End If
```

If the number of print jobs is greater than 0 (zero), this example displays a message in the status bar.




```vb
If Application.BackgroundPrintingStatus > 0 Then 
 StatusBar = Application.BackgroundPrintingStatus _ 
 &; " print jobs are queued up" 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

