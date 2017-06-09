---
title: TextStyles.Item Method (PowerPoint)
keywords: vbapp10.chm578003
f1_keywords:
- vbapp10.chm578003
ms.prod: powerpoint
api_name:
- PowerPoint.TextStyles.Item
ms.assetid: 3315d566-a46a-38cc-44b3-07c54ec3c6e5
ms.date: 06/08/2017
---


# TextStyles.Item Method (PowerPoint)

Returns a single text style from the specified  **[TextStyles](textstyles-object-powerpoint.md)** collection.


## Syntax

 _expression_. **Item**( **_Type_** )

 _expression_ A variable that represents a **TextStyles** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**PpTextStyleType**|The text style type.|

### Return Value

TextStyle


## Remarks

The  **Item** method is the default member for a collection. For example, the following two lines of code are equivalent:

 `ActivePresentation.Slides.Item(1)`

 `ActivePresentation.Slides(1)`

The  _Type_ parameter value can be one of these **PpTextStyleType** constants.


||
|:-----|
|**ppBodyStyle**|
|**ppDefaultStyle**|
|**ppTitleStyle**|

## See also


#### Concepts


[TextStyles Object](textstyles-object-powerpoint.md)

