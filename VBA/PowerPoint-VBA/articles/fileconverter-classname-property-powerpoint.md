---
title: FileConverter.ClassName Property (PowerPoint)
keywords: vbapp10.chm680004
f1_keywords:
- vbapp10.chm680004
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.ClassName
ms.assetid: dd024749-07e0-477c-2bba-5c78f2f222a6
ms.date: 06/08/2017
---


# FileConverter.ClassName Property (PowerPoint)

Returns a unique name that identifies the file converter. Read-only  **String**.


## Syntax

 _expression_. **ClassName**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[FileConverter Object](fileconverter-object-powerpoint.md)

