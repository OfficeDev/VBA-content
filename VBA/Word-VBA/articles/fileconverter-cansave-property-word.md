---
title: FileConverter.CanSave Property (Word)
keywords: vbawd10.chm161021956
f1_keywords:
- vbawd10.chm161021956
ms.prod: word
api_name:
- Word.FileConverter.CanSave
ms.assetid: a1de7523-5b9c-b606-4308-9445e3c4c76d
ms.date: 06/08/2017
---


# FileConverter.CanSave Property (Word)

 **True** if the specified file converter is designed to save files. Read-only **Boolean** .


## Syntax

 _expression_ . **CanSave**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

The  **[CanOpen](fileconverter-canopen-property-word.md)** property returns **True** if the specified file converter can be used to open (import) files.


## Example

This example determines whether the WordPerfect converter can be used to save files. If the return value is  **True** , the active document is saved in WordPerfect 6.x format.


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




```vb
If FileConverters(3).CanSave = True Then 
 MsgBox FileConverters(3).FormatName &; " can save files" 
Else 
 MsgBox FileConverters(3).FormatName &; " cannot save files" 
End If
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

