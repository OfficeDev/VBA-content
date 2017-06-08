---
title: Comment.Text Method (Excel)
keywords: vbaxl10.chm516076
f1_keywords:
- vbaxl10.chm516076
ms.prod: excel
api_name:
- Excel.Comment.Text
ms.assetid: 6a79c275-ba8e-799a-2e53-96347b1783a4
ms.date: 06/08/2017
---


# Comment.Text Method (Excel)

Sets comment text.


## Syntax

 _expression_ . **Text**( **_Text_** , **_Start_** , **_Overwrite_** )

 _expression_ A variable that represents a **Comment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|The text to be added.|
| _Start_|Optional| **Variant**|The character number where the added text will be placed. If this argument is omitted, any existing text in the comment is deleted.|
| _Overwrite_|Optional| **Variant**| **True** to overwrite the existing text. The default value is **False** (text is inserted).|

### Return Value

String


## See also


#### Concepts


[Comment Object](comment-object-excel.md)

