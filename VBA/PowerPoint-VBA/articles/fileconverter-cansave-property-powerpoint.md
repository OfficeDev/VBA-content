---
title: FileConverter.CanSave Property (PowerPoint)
keywords: vbapp10.chm680003
f1_keywords:
- vbapp10.chm680003
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.CanSave
ms.assetid: 64e1f21f-786e-8003-f99e-0dcb093af9d3
ms.date: 06/08/2017
---


# FileConverter.CanSave Property (PowerPoint)

 **True** if the specified file converter is designed to save files. Read-only **Boolean**.


## Syntax

 _expression_. **CanSave**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-powerpoint.md)** object.


## Remarks

The  **[CanOpen](fileconverter-canopen-property-powerpoint.md)** property returns **True** if the specified file converter can be used to open (import) files.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

This example determines whether the WordPerfect converter can be used to save files. If the return value is  **True**, the active document is saved in WordPerfect 6.x format.




```vb
Dim lngSaveFormat As Long

If Application.FileConverters("WordPerfect6x").CanSave = _
        True Then
		
    lngSaveFormat = _
        Application.FileConverters("WordPerfect6x").SaveFormat
		
    ActiveDocument.SaveAs FileName:="C:\Document.wp", _
        FileFormat:=lngSaveFormat

End If
```

This example displays a message that indicates whether the third converter in the FileConverters collection can save files.




```
If FileConverters(3).CanSave = True Then

    MsgBox FileConverters(3).FormatName &; " can save files"

Else

    MsgBox FileConverters(3).FormatName &; " cannot save files"

End If
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-powerpoint.md)

