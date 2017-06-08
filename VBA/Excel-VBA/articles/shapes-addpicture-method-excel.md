---
title: Shapes.AddPicture Method (Excel)
keywords: vbaxl10.chm638082
f1_keywords:
- vbaxl10.chm638082
ms.prod: excel
api_name:
- Excel.Shapes.AddPicture
ms.assetid: 50a46fce-e87d-d5a8-3218-7843788f82bb
ms.date: 06/08/2017
---


# Shapes.AddPicture Method (Excel)

Creates a picture from an existing file. Returns a  **Shape** object that represents the new picture.


## Syntax

 _expression_ . **AddPicture**( **_Filename_** , **_LinkToFile_** , **_SaveWithDocument_** , **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The file from which the OLE object is to be created.|
| _LinkToFile_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**| The file to link to.|
| _SaveWithDocument_|Required| **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)**|To save the picture with the document.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the picture relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the picture relative to the top of the document.|
| _Width_|Required| **Single**|The width of the picture, in points (enter -1 to retain the width of the existing file).|
| _Height_|Required| **Single**|The height of the picture, in points (enter -1 to retain the height of the existing file).|

### Return Value

Shape


## Remarks





| **MsoTriState** can be one of these **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants.|
| **msoCTrue**|
| **msoFalse** To make the picture an independent copy of the file.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** To link the picture to the file from which it was created.|


| **MsoTriState** can be one of these **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** constants.|
| **msoCTrue**|
| **msoFalse** To store only the link information in the document.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** To save the linked picture with the document into which it?s inserted. This argument must be **msoTrue** if _LinkToFile_ is **msoFalse** .|

## Example

This example adds a picture created from the file Music.bmp to  `myDocument`. The inserted picture is linked to the file from which it was created and is saved with  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddPicture _ 
    "c:\microsoft office\clipart\music.bmp", _ 
    True, True, 100, 100, 70, 70
```


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

