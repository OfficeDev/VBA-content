---
title: Application.ScreenRefresh Method (Word)
keywords: vbawd10.chm158335277
f1_keywords:
- vbawd10.chm158335277
ms.prod: word
api_name:
- Word.Application.ScreenRefresh
ms.assetid: 303db23c-492c-5e33-0363-7ef6433dc90e
ms.date: 06/08/2017
---


# Application.ScreenRefresh Method (Word)

Updates the display on the monitor with the current information in the video memory buffer.


## Syntax

 _expression_ . **ScreenRefresh**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Remarks

You can use this method after using the  **ScreenUpdating** property to disable screen updates. **ScreenRefresh** turns on screen updating for just one instruction and then immediately turns it off. Subsequent instructions don't update the screen until screen updating is turned on again with the **ScreenUpdating** property.


## Example

This example turns off screen updating, opens Test.doc, inserts text, refreshes the screen, and then closes the document (with changes saved).


```vb
Dim rngTemp As Range 
 
ScreenUpdating = False 
Documents.Open FileName:="C:\DOCS\TEST.DOC" 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
 
rngTemp.InsertBefore "new" 
Application.ScreenRefresh 
ActiveDocument.Close SaveChanges:=wdSaveChanges 
ScreenUpdating = True
```


## See also


#### Concepts


[Application Object](application-object-word.md)

