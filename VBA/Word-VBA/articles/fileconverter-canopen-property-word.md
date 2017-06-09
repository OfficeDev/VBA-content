---
title: FileConverter.CanOpen Property (Word)
keywords: vbawd10.chm161021957
f1_keywords:
- vbawd10.chm161021957
ms.prod: word
api_name:
- Word.FileConverter.CanOpen
ms.assetid: 0fe665dc-fe64-a61d-f6a5-a7ba2ff7b2d6
ms.date: 06/08/2017
---


# FileConverter.CanOpen Property (Word)

 **True** if the specified file converter is designed to open files. Read-only **Boolean** .


## Syntax

 _expression_ . **CanOpen**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

The  **[CanSave](fileconverter-cansave-property-word.md)** property returns **True** if the specified file converter can be used to save (export) files.


## Example

This example determines whether the first file converter is able to open files.


```vb
If FileConverters(1).CanOpen = True Then 
 MsgBox FileConverters(1).FormatName &; " can open files" 
End If
```

This example determines whether the WordPerfect6x file converter can be used to open files. If the CanOpen property returns True, a document named "Test.wp" is opened.




```vb
If FileConverters("WordPerfect6x").CanOpen = True Then 
 Documents.Open FileName:="C:\Test.wp", _ 
 Format:=FileConverters("WordPerfect6x").OpenFormat 
End If
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

