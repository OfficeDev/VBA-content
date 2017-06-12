---
title: EffectInformation.BuildByLevelEffect Property (PowerPoint)
keywords: vbapp10.chm655006
f1_keywords:
- vbapp10.chm655006
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.BuildByLevelEffect
ms.assetid: b839394f-1b58-4e12-9f55-38547cfd9bc1
ms.date: 06/08/2017
---


# EffectInformation.BuildByLevelEffect Property (PowerPoint)

Determines the level of the animation build effect. Read-only.


## Syntax

 _expression_. **BuildByLevelEffect**

 _expression_ A variable that represents a **EffectInformation** object.


### Return Value

MsoAnimateByLevel


## Remarks

The value returned by the  **BuildByLevelEffect** property can be one of these **MsoAnimateByLevel** constants.


||
|:-----|
|**msoAnimateChartAllAtOnce**|
|**msoAnimateChartByCategory**|
|**msoAnimateChartByCategoryElements**|
|**msoAnimateChartBySeries**|
|**msoAnimateChartBySeriesElements**|
|**msoAnimateDiagramAllAtOnce**|
|**msoAnimateDiagramBreadthByLevel**|
|**msoAnimateDiagramBreadthByNode**|
|**msoAnimateDiagramClockwise**|
|**msoAnimateDiagramClockwiseIn**|
|**msoAnimateDiagramClockwiseOut**|
|**msoAnimateDiagramCounterClockwise**|
|**msoAnimateDiagramCounterClockwiseIn**|
|**msoAnimateDiagramCounterClockwiseOut**|
|**msoAnimateDiagramDepthByBranch**|
|**msoAnimateDiagramDepthByNode**|
|**msoAnimateDiagramDown**|
|**msoAnimateDiagramInByRing**|
|**msoAnimateDiagramOutByRing**|
|**msoAnimateDiagramUp**|
|**msoAnimateLevelMixed**|
|**msoAnimateTextByAllLevels**|
|**msoAnimateTextByFifthLevel**|
|**msoAnimateTextByFirstLevel**|
|**msoAnimateTextByFourthLevel**|
|**msoAnimateTextBySecondLevel**|
|**msoAnimateTextByThirdLevel**|
|**msoAnimationLevelNone**|

## Example

The following example returns a build-by-level effect.


```vb
Sub QueryBuildByLevelEffect()

    Dim effMain As Effect

    Set effMain = ActivePresentation.Slides(1).TimeLine _
        .MainSequence(1)

    If effMain.EffectInformation.BuildByLevelEffect <> msoAnimateLevelNone Then
        ActivePresentation.Slides(1).TimeLine.MainSequence _
            .ConvertToTextUnitEffect Effect:=effMain, _
            UnitEffect:=msoAnimTextUnitEffectByWord
    End If

End Sub
```


## See also


#### Concepts



[EffectInformation Object](effectinformation-object-powerpoint.md)

