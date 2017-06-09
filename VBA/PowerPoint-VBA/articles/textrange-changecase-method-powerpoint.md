---
title: TextRange.ChangeCase Method (PowerPoint)
keywords: vbapp10.chm569031
f1_keywords:
- vbapp10.chm569031
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.ChangeCase
ms.assetid: a14edb26-7ec3-5fb5-7590-cd67a75c1f03
ms.date: 06/08/2017
---


# TextRange.ChangeCase Method (PowerPoint)

Changes the case of the specified text.


## Syntax

 _expression_. **ChangeCase**( **_Type_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[PpChangeCase](ppchangecase-enumeration-powerpoint.md)**|Specifies the way the case will be changed.|

## Example

This example sets title case capitalization for the title on slide one in the the active presentation.


```vb
Application.ActivePresentation.Slides(1).Shapes.Title.TextFrame _
    .TextRange.ChangeCase ppCaseTitle
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

