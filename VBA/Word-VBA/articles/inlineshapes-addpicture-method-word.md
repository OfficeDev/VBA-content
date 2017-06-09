---
title: InlineShapes.AddPicture Method (Word)
keywords: vbawd10.chm162070628
f1_keywords:
- vbawd10.chm162070628
ms.prod: word
api_name:
- Word.InlineShapes.AddPicture
ms.assetid: 89c5f587-d591-d56b-d52a-fd21073f76fb
ms.date: 06/08/2017
---


# InlineShapes.AddPicture Method (Word)

Adds a picture to a document. Returns an  **[InlineShape](inlineshape-object-word.md)** object that represents the picture.


## Syntax

 _expression_ . **AddPicture**( **_FileName_** , **_LinkToFile_** , **_SaveWithDocument_** , **_Range_** )

 _expression_ Required. A variable that represents an **[InlineShapes](inlineshapes-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path and file name of the picture.|
| _LinkToFile_|Optional| **Variant**| **True** to link the picture to the file from which it was created. **False** to make the picture an independent copy of the file. The default value is **False** .|
| _SaveWithDocument_|Optional| **Variant**| **True** to save the linked picture with the document. The default value is **False** .|
| _Range_|Optional| **Variant**|The location where the picture will be placed in the text. If the range isn't collapsed, the picture replaces the range; otherwise, the picture is inserted. If this argument is omitted, the picture is placed automatically.|

## See also


#### Concepts


[InlineShapes Collection Object](inlineshapes-object-word.md)

