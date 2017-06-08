---
title: FileConverter.FormatName Property (Word)
keywords: vbawd10.chm161021952
f1_keywords:
- vbawd10.chm161021952
ms.prod: word
api_name:
- Word.FileConverter.FormatName
ms.assetid: bfae89b4-14dd-ed73-6174-52c6cc7a9017
ms.date: 06/08/2017
---


# FileConverter.FormatName Property (Word)

Returns the name of the specified file converter. Read-only  **String** .


## Syntax

 _expression_ . **FormatName**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-word.md)** object.


## Remarks

The format names appear in the  **Save as type** box in the **Save As** dialog box ( **File** menu).


## Example

This example displays the format name of the first converter in the FileConverters collection.


```vb
MsgBox FileConverters(1).FormatName
```

This example uses the AvailableConv() array to store the names of all the available file converters.




```vb
Dim intTemp As Integer 
Dim fcLoop As FileConverter 
Dim AvailableConv As Variant 
 
ReDim AvailableConv(FileConverters.Count - 1) 
 
intTemp = 0 
 
For Each fcLoop In FileConverters 
 AvailableConv(intTemp) = fcLoop.FormatName 
 intTemp = intTemp + 1 
Next fcLoop
```


## See also


#### Concepts


[FileConverter Object](fileconverter-object-word.md)

