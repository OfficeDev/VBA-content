---
title: Options.ShowFormatError Property (Word)
keywords: vbawd10.chm162988480
f1_keywords:
- vbawd10.chm162988480
ms.prod: word
api_name:
- Word.Options.ShowFormatError
ms.assetid: 619ccdb4-020c-d6c7-48a8-2d2e56377577
ms.date: 06/08/2017
---


# Options.ShowFormatError Property (Word)

 **True** for Microsoft Word to mark inconsistencies in formatting by placing a squiggly underline beneath text formatted similarly to other formatting that is used more frequently in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowFormatError**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Example

This example enables Word to keep track of formatting in documents but does not display a squiggly underline beneath text.


```vb
Sub ShowFormatErrors() 
 
 With Options 
 .FormatScanning = True 'Enables keeping track of formatting 
 .ShowFormatError = False 
 End With 
 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

