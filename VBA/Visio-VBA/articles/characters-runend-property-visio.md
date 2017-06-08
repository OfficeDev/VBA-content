---
title: Characters.RunEnd Property (Visio)
keywords: vis_sdr.chm10214280
f1_keywords:
- vis_sdr.chm10214280
ms.prod: visio
api_name:
- Visio.Characters.RunEnd
ms.assetid: 4c9d0f81-8b6d-d5c3-98a1-1d0b39f8193a
ms.date: 06/08/2017
---


# Characters.RunEnd Property (Visio)

Returns the ending index of a type of run?a sequence of characters that share a particular attribute, such as character, paragraph, or tab formatting; or a word, paragraph, or field. Read-only.


## Syntax

 _expression_ . **RunEnd**( **_RunType_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _RunType_|Required| **Integer**|The type of run to get.|

### Return Value

Long


## Remarks

In a ShapeSheet window, each row in the Character and Paragraph sections represents a run of the corresponding format in a shape's text. Certain words may be bold or italic, or one paragraph may be centered and another left-aligned. Each change of format represents a run of that format. Similarly, delimiters such as spaces and paragraph marks represent the beginning and end of words, paragraphs, and fields.

In addition, you can retrieve rows that represent runs of character, paragraph, and tab formats by specifying a row index as an argument to the  **CellsSRC** property of a shape.

Use the  **RunEnd** property to determine the end of a sequence of identically formatted characters or the end of a word, paragraph, or field. You can get the **IsField** property to determine whether a run is a field.

The index that the  **RunEnd** property returns is greater than or equal to the ending index of a **Characters** object. If the **End** property of the **Characters** object is already at the end of a run, the value of the **RunEnd** property is equal to the value of the **End** property.

Use the  _RunType_ argument to specify the type of run you want. You can use any of the constants declared by the Visio type library in **[VisRunTypes Constants](visruntypes-enumeration-visio.md)** . To find a list of _RunType_ values, see the **RunBegin** property.


