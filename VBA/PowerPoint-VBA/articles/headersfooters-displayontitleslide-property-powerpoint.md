---
title: HeadersFooters.DisplayOnTitleSlide Property (PowerPoint)
keywords: vbapp10.chm542007
f1_keywords:
- vbapp10.chm542007
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters.DisplayOnTitleSlide
ms.assetid: adcf0504-50ce-b302-c61f-08425acf847c
ms.date: 06/08/2017
---


# HeadersFooters.DisplayOnTitleSlide Property (PowerPoint)

Determines whether the footer, date and time, and slide number appear on the title slide. Applies to slide masters. Read/write. 


## Syntax

 _expression_. **DisplayOnTitleSlide**

 _expression_ A variable that represents a **HeadersFooters** object.


### Return Value

MsoTriState


## Remarks

The value of the  **DisplayOnTitleSlide** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The footer, date and time, and slide number appears on all slides except the title slide.|
|**msoTrue**| The footer, date and time, and slide number appear on the title slide.|

## Example

This example sets the footer, date and time, and slide number to not appear on the title slide.


```vb
Application.ActivePresentation.SlideMaster.HeadersFooters.DisplayOnTitleSlide = msoFalse
```


## See also


#### Concepts


[HeadersFooters Object](headersfooters-object-powerpoint.md)

