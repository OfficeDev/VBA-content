---
title: FileConverter.Extensions Property (PowerPoint)
keywords: vbapp10.chm680006
f1_keywords:
- vbapp10.chm680006
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.Extensions
ms.assetid: 4003e78b-c931-94a4-e53a-3bedb9512a6a
ms.date: 06/08/2017
---


# FileConverter.Extensions Property (PowerPoint)

Returns the file name extensions associated with the specified  **FileConverter** object. Read-only **String**.


## Syntax

 _expression_. **Extensions**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

This example displays the name and file name extensions for the first file converter.




```vb
Dim fcTemp As FileConverter



Set fcTemp = FileConverters(1)

MsgBox "The file name extensions for " &; fcTemp.FormatName _
    &; " files are: " &; fcTemp.Extensions
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-powerpoint.md)

