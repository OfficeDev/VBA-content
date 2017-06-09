---
title: FileConverter.Extensions Property (Word)
keywords: vbawd10.chm161021960
f1_keywords:
- vbawd10.chm161021960
ms.prod: word
api_name:
- Word.FileConverter.Extensions
ms.assetid: 18a9819b-ddc3-5928-8ce7-882d00d3f5c9
ms.date: 06/08/2017
---


# FileConverter.Extensions Property (Word)

Returns the file name extensions associated with the specified  **FileConverter** object. Read-only **String** .


## Syntax

 _expression_ . **Extensions**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Example

This example displays the name and file name extensions for first file converter.


```vb
Dim fcTemp As FileConverter 
 
Set fcTemp = FileConverters(1) 
MsgBox "The file name extensions for " &; fcTemp.FormatName _ 
 &; " files are: " &; fcTemp.Extensions
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

