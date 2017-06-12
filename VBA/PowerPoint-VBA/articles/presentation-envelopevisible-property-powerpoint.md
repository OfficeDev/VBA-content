---
title: Presentation.EnvelopeVisible Property (PowerPoint)
keywords: vbapp10.chm583057
f1_keywords:
- vbapp10.chm583057
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.EnvelopeVisible
ms.assetid: e2a58d05-df9b-0fc6-a1d4-3349b7efa111
ms.date: 06/08/2017
---


# Presentation.EnvelopeVisible Property (PowerPoint)

Determines whether the e-mail message header is visible in the document window. Read/write.


## Syntax

 _expression_. **EnvelopeVisible**

 _expression_ A variable that represents an **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **EnvelopeVisible** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| The e-mail message header is not visible in the document window. The default.|
|**msoTrue**| The e-mail message header is visible in the document window.|

## Example

This example displays the e-mail message header.


```vb
ActivePresentation.EnvelopeVisible = msoTrue
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

