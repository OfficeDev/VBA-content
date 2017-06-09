---
title: Options.FormatScanning Property (Word)
keywords: vbawd10.chm162988481
f1_keywords:
- vbawd10.chm162988481
ms.prod: word
api_name:
- Word.Options.FormatScanning
ms.assetid: 7557b88e-2f16-47e9-cc3b-05019dba9896
ms.date: 06/08/2017
---


# Options.FormatScanning Property (Word)

 **True** for Microsoft Word to keep track of all formatting in a document. Read/write **Boolean** .


## Syntax

 _expression_ . **FormatScanning**

 _expression_ An expression that returns an **[Options](options-object-word.md)** object.


## Remarks

Enabling the  **FormatScanning** property allows you to identify all unique formatting in your document, so you can easily apply the same formatting to new text and quickly replace or modify all instances of a given formatting within a document.


## Example

This example enables Word to keep track of formatting in documents but disables displaying a squiggly underline beneath text formatted similarly to other formatting that is used more frequently in a document.


```vb
Sub ShowFormatErrors() 
 With Options 
 .FormatScanning = True 
 .ShowFormatError = False 'Disables displaying squiggly underline 
 End With 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

