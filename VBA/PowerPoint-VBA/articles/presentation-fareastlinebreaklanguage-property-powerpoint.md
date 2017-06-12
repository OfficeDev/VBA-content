---
title: Presentation.FarEastLineBreakLanguage Property (PowerPoint)
keywords: vbapp10.chm583048
f1_keywords:
- vbapp10.chm583048
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.FarEastLineBreakLanguage
ms.assetid: e0acc33d-0cb0-5422-4238-26b4071fb48c
ms.date: 06/08/2017
---


# Presentation.FarEastLineBreakLanguage Property (PowerPoint)

Returns or sets the language used to determine which line break level is used when the line break control option is turned on. Read/write.


## Syntax

 _expression_. **FarEastLineBreakLanguage**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoFarEastLineBreakLanguageID


## Remarks

The value of the  **FarEastLineBreakLanguage** property can be one of these **MsoFarEastLineBreakLanguageID** constants.


||
|:-----|
|**MsoFarEastLineBreakLanguageJapanese**|
|**MsoFarEastLineBreakLanguageKorean**|
|**MsoFarEastLineBreakLanguageSimplifiedChinese**|
|**MsoFarEastLineBreakLanguageTraditionalChinese**|

## Example

The following example sets the line break language to Japanese.


```vb
ActivePresentation.FarEastLineBreakLanguage =  MsoFarEastLineBreakLanguageJapanese
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

