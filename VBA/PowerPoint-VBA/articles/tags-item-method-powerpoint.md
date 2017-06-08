---
title: Tags.Item Method (PowerPoint)
keywords: vbapp10.chm611003
f1_keywords:
- vbapp10.chm611003
ms.prod: powerpoint
api_name:
- PowerPoint.Tags.Item
ms.assetid: 66e4b84b-4bcc-d526-fa69-0ecfc52ef649
ms.date: 06/08/2017
---


# Tags.Item Method (PowerPoint)

Returns a single tag from the specified  **[Tags](tags-object-powerpoint.md)** collection.


## Syntax

 _expression_. **Item**( **_Name_** )

 _expression_ A variable that represents a **Tags** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the single tag in the collection to be returned.|

### Return Value

String


## Example

This example hides all slides in the active presentation that don't have the value "east" for the "region" tag.


```vb
For Each s In ActivePresentation.Slides

    If s.Tags.Item("region") <> "east" Then

        s.SlideShowTransition.Hidden = True

    End If

Next
```


## See also


#### Concepts


[Tags Object](tags-object-powerpoint.md)

