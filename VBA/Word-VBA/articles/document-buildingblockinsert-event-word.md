---
title: Document.BuildingBlockInsert Event (Word)
keywords: vbawd10.chm4001016
f1_keywords:
- vbawd10.chm4001016
ms.prod: word
api_name:
- Word.Document.BuildingBlockInsert
ms.assetid: 6c4b1f1f-da22-63b9-a3d9-21d7bedb4b5b
ms.date: 06/08/2017
---


# Document.BuildingBlockInsert Event (Word)

Occurs when you insert a building block into a document. .


## Syntax

Private Sub  _expression_ _**BuildingBlockInsert**( **_Range_** , **_Name_** , **_Category_** , **_Type_** , **_Template_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|Specifies the position where the building block is inserted.|
| _Name_|Required| **String**|Specifies the name of the building block.|
| _Category_|Required| **String**|Specifies the building block category.|
| _Type_|Required| **String**|Specifies the type of building block.|
| _Template_|Required| **String**|Specifies the name of the template that contains the building block.|

## Remarks

For information about using events with a  **Document** object, see[Using Events with the Document Object](http://msdn.microsoft.com/library/2b043342-436a-5421-e8af-3c2c49684960%28Office.15%29.aspx).


## See also


#### Concepts


[Document Object](document-object-word.md)

