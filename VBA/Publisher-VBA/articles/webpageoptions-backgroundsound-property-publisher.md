---
title: WebPageOptions.BackgroundSound Property (Publisher)
keywords: vbapb10.chm544774
f1_keywords:
- vbapb10.chm544774
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.BackgroundSound
ms.assetid: c6be30e0-28ea-e269-c546-48e0eb284ac4
ms.date: 06/08/2017
---


# WebPageOptions.BackgroundSound Property (Publisher)

Returns or sets a  **String** that specifies the path to a sound file that is played when the Web page is loaded in a Web browser. Read/write.


## Syntax

 _expression_. **BackgroundSound**

 _expression_A variable that represents a  **WebPageOptions** object.


### Return Value

String


## Remarks

The path to the background sound file must be a network path or a local path; an http:// address will not work.

If  **BackgroundSound** is specified, the background sound will play once by default. The **[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)** method can be used to specify whether the background sound should be played infinitely, and if it should not, to specify the number of times the background sound file should be played.

The background sound can be any of the following file types:



|*.wav|
|*.mid|
|*.midi|
|*.rmi|
|*.au|
|*.aif|
|*.aiff|

## Example

The following example sets the background sound for page four of the active Web publication to a .wav file on the local computer. This .wav file will play once when the page is loaded in a Web browser.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
End With
```


