---
title: LinkFormat.SourceFullName Property (Word)
keywords: vbawd10.chm154206229
f1_keywords:
- vbawd10.chm154206229
ms.prod: word
api_name:
- Word.LinkFormat.SourceFullName
ms.assetid: a55a6834-3325-567c-47da-76e976bc6ebf
ms.date: 06/08/2017
---


# LinkFormat.SourceFullName Property (Word)

Returns or sets the path and name of the source file for the specified linked OLE object, picture, or field. Read/write  **String** .


## Syntax

 _expression_ . **SourceFullName**

 _expression_ An expression that returns a **[LinkFormat](linkformat-object-word.md)** object.


## Remarks

Using this property is equivalent to using in sequence the  **[SourcePath](linkformat-sourcepath-property-word.md)** , **[PathSeparator](application-pathseparator-property-word.md)** , and **[SourceName](linkformat-sourcename-property-word.md)** properties.


## Example

This example sets MyExcel.xls as the source file for shape one on the active document and specifies that the OLE object be updated automatically.


```vb
With ActiveDocument.Shapes(1) 
 If .Type = msoLinkedOLEObject Then 
 With .LinkFormat 
 .SourceFullName = "c:\my documents\myExcel.xls" 
 .AutoUpdate = True 
 End With 
 End If 
End With
```


## See also


#### Concepts


[LinkFormat Object](linkformat-object-word.md)

