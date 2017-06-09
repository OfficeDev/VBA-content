---
title: Presentations.Add Method (PowerPoint)
keywords: vbapp10.chm522004
f1_keywords:
- vbapp10.chm522004
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.Add
ms.assetid: 9a09ad9b-c52d-9fd6-20ef-68b694596ed2
ms.date: 06/08/2017
---


# Presentations.Add Method (PowerPoint)

Creates a presentation. Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the new presentation.


## Syntax

 _expression_. **Add**( **_WithWindow_** )

 _expression_ A variable that represents a **Presentations** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _WithWindow_|Optional|**MsoTriState**|Whether the presentation appears in a visible window.|

### Return Value

Presentation


## Remarks

The  _WithWindow_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The new presentation isn't visible.|
|**msoTrue**|The default. Creates the presentation in a visible window.|

## Example

This example creates a presentation, adds a slide to it, and then saves the presentation.


```vb
With Presentations.Add

    .Slides.Add Index:=1, Layout:=ppLayoutTitle

    .SaveAs "Sample"

End With


```


## See also


#### Concepts


[Presentations Object](presentations-object-powerpoint.md)

