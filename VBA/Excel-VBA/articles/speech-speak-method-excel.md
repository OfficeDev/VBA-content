---
title: Speech.Speak Method (Excel)
keywords: vbaxl10.chm718073
f1_keywords:
- vbaxl10.chm718073
ms.prod: excel
api_name:
- Excel.Speech.Speak
ms.assetid: d17dcf63-c837-a5b5-8267-44767b38700a
ms.date: 06/08/2017
---


# Speech.Speak Method (Excel)

Microsoft Excel plays back the text string that is passed as an argument.


## Syntax

 _expression_ . **Speak**( **_Text_** , **_SpeakAsync_** , **_SpeakXML_** , **_Purge_** )

 _expression_ A variable that represents a **Speech** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text to be spoken.|
| _SpeakAsync_|Optional| **Variant**| **True** will cause the _Text_ to be spoken asynchronously (the method will not wait of the Text to be spoken). **False** will cause the _Text_ to be spoken synchronously (the method waits for the _Text_ to be spoken before continuing). The default is **False** .|
| _SpeakXML_|Optional| **Variant**| **True** will cause the _Text_ to be interpreted as XML. **False** will cause the _Text_ to not be interpreted as XML, so any XML tags will be read and not interpreted. The default is **False** .|
| _Purge_|Optional| **Variant**| **True** will cause current speech to be terminated and any buffered text to be purged before _Text_ is spoken. **False** will not cause the current speech to be terminated and will not purge the buffered text before _Text_ is spoken. The default is **False** .|

## Example

In this example, Microsoft Excel speaks "Hello".


```vb
Sub UseSpeech() 
 
 Application.Speech.Speak "Hello" 
 
End Sub
```


## See also


#### Concepts


[Speech Object](speech-object-excel.md)

