---
title: WebPageOptions.BackgroundSoundLoopForever Property (Publisher)
keywords: vbapb10.chm544775
f1_keywords:
- vbapb10.chm544775
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.BackgroundSoundLoopForever
ms.assetid: f2e90665-09e9-5215-59e4-f93e4469d0df
ms.date: 06/08/2017
---


# WebPageOptions.BackgroundSoundLoopForever Property (Publisher)

Returns a  **Boolean** value that specifies whether the background sound attached to the Web page should be repeated infinitely. Read-only.


## Syntax

 _expression_. **BackgroundSoundLoopForever**

 _expression_A variable that represents a  **WebPageOptions** object.


### Return Value

Boolean


## Remarks

The  **[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)** method is used to specify whether the background sound should be repeated infinitely after the page is loaded. Until the **SetBackgroundSoundRepeat** method is used to specify whether the background sound should be played infinitely, **BackgroundSoundLoopForever** is **False**.


## Example

The following example sets the background sound for page four of the active Web publication to a .wav file on the local computer. If  **BackgroundSoundLoopForever** is **False**, the  **SetBackgroundSoundRepeat** method is used to specify that the background sound should be repeated infinitely. The **BackgroundSoundLoopForever** property will now be **True**.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopForever = False Then 
 .SetBackgroundSoundRepeat RepeatForever:=True 
 End If 
End With
```


