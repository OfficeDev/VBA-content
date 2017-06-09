---
title: Characters.RunBegin Property (Visio)
keywords: vis_sdr.chm10214275
f1_keywords:
- vis_sdr.chm10214275
ms.prod: visio
api_name:
- Visio.Characters.RunBegin
ms.assetid: 6397f797-c481-e2f0-ec38-61a799762552
ms.date: 06/08/2017
---


# Characters.RunBegin Property (Visio)

Returns the beginning index of a type of run?a sequence of characters that share a particular attribute, such as character, paragraph, or tab formatting; or a word, paragraph, or field. Read-only.


## Syntax

 _expression_ . **RunBegin**( **_RunType_** )

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

Use the  **RunBegin** property to determine the beginning of a sequence of identically formatted characters or the beginning of a word, paragraph, or field. You can get the **IsField** property to determine whether a run is a field.

The index that the  **RunBegin** property returns is less than or equal to the beginning index of a **Characters** object. If the **Begin** property of the **Characters** object is already at the start of a run, the value of the **RunBegin** property is equal to the value of **Begin** .

Use the  _RunType_ argument to specify the type of run you want. You can use any of the following constants declared by the Visio type library in member **VisRunTypes** .



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visCharPropRow**|1 |Reports runs of characters with common character properties. Corresponds to the set of characters covered by one row in the shape's Character section. |
| **visParaPropRow**|2 |Reports runs of characters with common paragraph properties. Corresponds to the set of characters covered by one row in the shape's Paragraph section. |
| **visTabPropRow**|3 |Reports runs of characters with common tab properties. Corresponds to the set of characters covered by one row in the shape's Tabs section. |
| **visWordRun**|10 |Reports runs whose boundaries are between successive words in the shape's text. Mimics double-clicking to select text. |
| **visParaRun**|11 |Reports runs whose boundaries are between successive paragraphs in the shape's text. Mimics triple-clicking to select text. |
| **visFieldRun**|20 |Reports runs whose boundaries are between characters that are and are not the result of the expansion of a text field, or between characters that are the result of the expansion of distinct text fields. |

