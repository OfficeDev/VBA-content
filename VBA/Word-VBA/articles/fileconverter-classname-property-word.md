---
title: FileConverter.ClassName Property (Word)
keywords: vbawd10.chm161021953
f1_keywords:
- vbawd10.chm161021953
ms.prod: word
api_name:
- Word.FileConverter.ClassName
ms.assetid: 71124adf-11fc-e42d-a9f5-940f7fea97af
ms.date: 06/08/2017
---


# FileConverter.ClassName Property (Word)

Returns a unique name that identifies the file converter. Read-only  **String** .


## Syntax

 _expression_ . **ClassName**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Example

This example displays the class name and format name of the first converter in the FileConverters collection.


```vb
MsgBox "ClassName= " &; FileConverters(1).ClassName &; vbCr _ 
 &; "FormatName= " &; FileConverters(1).FormatName
```

If an HTML file converter is available, this example sets the HTML format as the default save format.




```vb
Dim fcLoop As FileConverter 
 
For Each fcLoop In FileConverters 
 If fcLoop.ClassName = "HTML" Then _ 
 Application.DefaultSaveFormat = "HTML" 
Next fcLoop
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

