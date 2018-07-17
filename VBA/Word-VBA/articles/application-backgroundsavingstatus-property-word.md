---
title: Application.BackgroundSavingStatus Property (Word)
keywords: vbawd10.chm158335061
f1_keywords:
- vbawd10.chm158335061
ms.prod: word
api_name:
- Word.Application.BackgroundSavingStatus
ms.assetid: 9cf29eb4-fc80-91ad-2867-6dc9d48e11c7
ms.date: 06/08/2017
---


# Application.BackgroundSavingStatus Property (Word)

Returns the number of files queued up to be saved in the background. Read-only  **Long** .


## Syntax

 _expression_ . **BackgroundSavingStatus**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays in the status bar the number of documents currently being saved.


```vb
Options.BackgroundSave = True 
Documents.Add 
ActiveDocument.SaveAs 
 While Application.BackgroundSavingStatus <> 0 
 StatusBar = "Documents remaining to save: " _ 
 &; Application.BackgroundSavingStatus 
 DoEvents 
Wend
```


## See also


#### Concepts


[Application Object](application-object-word.md)

