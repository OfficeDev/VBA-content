---
title: AnimationSettings.TextUnitEffect Property (PowerPoint)
keywords: vbapp10.chm565012
f1_keywords:
- vbapp10.chm565012
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationSettings.TextUnitEffect
ms.assetid: 6948db54-775a-39d6-9d90-99ad25f9cb80
ms.date: 06/08/2017
---


# AnimationSettings.TextUnitEffect Property (PowerPoint)

Indicates whether the text in the specified shape is animated paragraph by paragraph, word by word, or letter by letter. Read/write.


## Syntax

 _expression_. **TextUnitEffect**

 _expression_ A variable that represents an **AnimationSettings** object.


### Return Value

PpTextUnitEffect


## Remarks

The value of the  **TextUnitEffect** property can be one of these **PpTextUnitEffect** constants.


||
|:-----|
|**ppAnimateByCharacter**|
|**ppAnimateByParagraph**|
|**ppAnimateByWord**|
|**ppAnimateUnitMixed**|
For the  **TextUnitEffect** property setting to take effect, the **[TextLevelEffect](animationsettings-textleveleffect-property-powerpoint.md)** property for the specified shape must have a value other than **ppAnimateLevelNone** or **ppAnimateByAllLevels**, and the **[Animate](animationsettings-animate-property-powerpoint.md)** property must be set to **True**.


## Example

This example adds a title slide and title text to the active presentation and sets the title to be built letter by letter.


```vb
With ActivePresentation.Slides.Add(Index:=1, _
    Layout:=ppLayoutTitleOnly).Shapes(1)

    .TextFrame.TextRange.Text = "Sample title"
    With .AnimationSettings
        .Animate = True
        .TextLevelEffect = ppAnimateByFirstLevel
        .TextUnitEffect = ppAnimateByCharacter
        .EntryEffect = ppEffectFlyFromLeft
    End With

End With
```


## See also


#### Concepts


[AnimationSettings Object](animationsettings-object-powerpoint.md)

