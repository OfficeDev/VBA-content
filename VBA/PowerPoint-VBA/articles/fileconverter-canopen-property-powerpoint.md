---
title: FileConverter.CanOpen Property (PowerPoint)
keywords: vbapp10.chm680002
f1_keywords:
- vbapp10.chm680002
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.CanOpen
ms.assetid: 9a5a2fea-0f09-9dfe-c75a-e8811d53c27f
ms.date: 06/08/2017
---


# FileConverter.CanOpen Property (PowerPoint)

 **True** if the specified file converter is designed to open files. Read-only **Boolean**.


## Syntax

 _expression_. **CanOpen**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-powerpoint.md)** object.


## Remarks

The  **[CanSave](fileconverter-cansave-property-powerpoint.md)** property returns **True** if the specified file converter can be used to save (export) files.


## Example

This example determines whether the first file converter is able to open files.


```
If FileConverters(1).CanOpen = True Then

    MsgBox FileConverters(1).FormatName &; " can open files"

End If
```




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

This example determines whether the WordPerfect6x file converter can be used to open files. If the CanOpen property returns True, a document named "Test.wp" is opened.




```
If FileConverters("WordPerfect6x").CanOpen = True Then
    Documents.Open FileName:="C:\Test.wp", _
        Format:=FileConverters("WordPerfect6x").OpenFormat
End If
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-powerpoint.md)

