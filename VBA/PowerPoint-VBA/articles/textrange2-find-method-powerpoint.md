---
title: TextRange2.Find Method (PowerPoint)
ms.assetid: 6d7d1ef8-8a61-4fbd-b157-22f64e6f8a6f
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.Find Method (PowerPoint)

Searches a  **TextRange2** object for a subset of text.


## Syntax

 _expression_. **Find**( **_FindWhat_**, **_After_**, **_MatchCase_**, **_WholeWords_** )

 _expression_ An expression that returns a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FindWhat_|Required|**String**|Contains the text to find.|
| _After_|Optional|**Long**|Specifies the point in the text range to start the search.|
| _MatchCase_|Optional|**MsoTriState**|Specifies if the target text must exactly match the case of the search text. |
| _WholeWords_|Optional|**MsoTriState**|Specifies that only whole words will be searched.|

### Return Value

TextRange2


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


