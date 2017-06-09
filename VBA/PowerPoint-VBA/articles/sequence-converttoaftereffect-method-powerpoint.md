---
title: Sequence.ConvertToAfterEffect Method (PowerPoint)
keywords: vbapp10.chm651009
f1_keywords:
- vbapp10.chm651009
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.ConvertToAfterEffect
ms.assetid: bbd340a5-d0c4-1db9-259c-ee43c079309a
ms.date: 06/08/2017
---


# Sequence.ConvertToAfterEffect Method (PowerPoint)

Specifies what an effect should do after it is finished. Returns an  **[Effect](effect-object-powerpoint.md)** object that represents an after effect.


## Syntax

 _expression_. **ConvertToAfterEffect**( **_Effect_**, **_After_**, **_DimColor_**, **_DimSchemeColor_** )

 _expression_ A variable that represents a **Sequence** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**|The effect to which the after effect will be added.|
| _After_|Required|**[MsoAnimAfterEffect](msoanimaftereffect-enumeration-powerpoint.md)**|The behavior of the after effect.|
| _DimColor_|Optional|**MsoRGBType**|A single color to apply the after effect.|
| _DimSchemeColor_|Optional|**[PpColorSchemeIndex](ppcolorschemeindex-enumeration-powerpoint.md)**|A predefined color scheme to apply to the after effect.|

### Return Value

Effect


## Remarks

Do not use both the DimColor and DimSchemeColor arguments in the same call to this method. An after effect may have one color, or use a predefined color scheme, but not both.


## Example

The following example sets a dim color for an after effect on the first shape on the first slide in the active presentation. This example assume there is a shape on the first slide.


```vb
Sub ConvertToDim()

    Dim shpSelected As Shape
    Dim sldActive As Slide
    Dim effConvert As Effect

    Set sldActive = ActivePresentation.Slides(1)
    Set shpSelected = sldActive.Shapes(1)

    ' Add an animation effect.
    Set effConvert = sldActive.TimeLine.MainSequence.AddEffect _
        (Shape:=shpSelected, effectId:=msoAnimEffectBounce)

    ' Add a dim after effect.
    Set effConvert = sldActive.TimeLine.MainSequence.ConvertToAfterEffect _
        (Effect:=effConvert, After:=msoAnimAfterEffectDim, _
        DimColor:=RGB(Red:=255, Green:=255, Blue:=255))

End Sub
```


## See also


#### Concepts


[Sequence Object](sequence-object-powerpoint.md)

