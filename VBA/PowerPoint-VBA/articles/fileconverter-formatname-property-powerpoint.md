---
title: FileConverter.FormatName Property (PowerPoint)
keywords: vbapp10.chm680007
f1_keywords:
- vbapp10.chm680007
ms.prod: powerpoint
api_name:
- PowerPoint.FileConverter.FormatName
ms.assetid: 50d92230-05a5-7dc1-115c-0e32ba0a76f3
ms.date: 06/08/2017
---


# FileConverter.FormatName Property (PowerPoint)

Returns the name of the specified file converter. Read-only  **String**.


## Syntax

 _expression_. **FormatName**

 _expression_ A variable that represents a **[FileConverter](fileconverter-object-powerpoint.md)** object.


## Remarks

The format names appear in the  **Save as type** box in the **Save As** dialog box.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[FileConverter Object](fileconverter-object-powerpoint.md)

