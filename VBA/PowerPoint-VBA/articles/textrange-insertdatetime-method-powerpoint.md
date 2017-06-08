---
title: TextRange.InsertDateTime Method (PowerPoint)
keywords: vbapp10.chm569020
f1_keywords:
- vbapp10.chm569020
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.InsertDateTime
ms.assetid: b1f6c2db-2524-f76e-eee2-8f177b08dcde
ms.date: 06/08/2017
---


# TextRange.InsertDateTime Method (PowerPoint)

Inserts the date and time in the specified text range. Returns a  **TextRange** object that represents the inserted text.


## Syntax

 _expression_. **InsertDateTime**( **_DateTimeFormat_**, **_InsertAsField_** )

 _expression_ A variable that represents a **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DateTimeFormat_|Required|**PpDateTimeFormat**|A format for the date and time.|
| _InsertAsField_|Optional|**MsoTriState**|Determines whether the inserted date and time will be updated each time the presentation is opened.|

### Return Value

TextRange


## Remarks

The  _DateTimeFormat_ parameter value can be one of these **PpDateTimeFormat** constants.


||
|:-----|
|**ppDateTimeddddMMMMddyyyy**|
|**ppDateTimedMMMMyyyy**|
|**ppDateTimedMMMyy**|
|**ppDateTimeFormatMixed**|
|**ppDateTimeHmm**|
|**ppDateTimehmmAMPM**|
|**ppDateTimeHmmss**|
|**ppDateTimehmmssAMPM**|
|**ppDateTimeMdyy**|
|**ppDateTimeMMddyyHmm**|
|**ppDateTimeMMddyyhmmAMPM**|
|**ppDateTimeMMMMdyyyy**|
|**ppDateTimeMMMMyy**|
|**ppDateTimeMMyy**|
The  _InsertAsField_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default.|
|**msoTrue**|Updates the inserted date and time each time the presentation is opened.|

## Example

This example inserts the date and time after the first sentence of the first paragraph in shape two on slide one in the active presentation.


```vb
Set sh = Application.ActivePresentation.Slides(1).Shapes(2)

Set sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)

sentOne.InsertAfter.InsertDateTime ppDateTimeMdyy
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

