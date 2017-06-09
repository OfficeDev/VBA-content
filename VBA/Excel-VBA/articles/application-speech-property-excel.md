---
title: Application.Speech Property (Excel)
keywords: vbaxl10.chm133285
f1_keywords:
- vbaxl10.chm133285
ms.prod: excel
api_name:
- Excel.Application.Speech
ms.assetid: 981d5eef-55ff-54ee-a3ca-f009a6a575da
ms.date: 06/08/2017
---


# Application.Speech Property (Excel)

Returns a  **[Speech](speech-object-excel.md)** object.


## Syntax

 _expression_ . **Speech**

 _expression_ A variable that represents an **Application** object.


## Example

In the following example, Microsoft Excel plays back "Hello". This example assumes speech features have been installed on the host system.


```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

