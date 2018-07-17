---
title: Effect Object (PowerPoint)
keywords: vbapp10.chm652000
f1_keywords:
- vbapp10.chm652000
ms.prod: powerpoint
api_name:
- PowerPoint.Effect
ms.assetid: 359ac3da-86cd-8003-d691-349d20fd1777
ms.date: 06/08/2017
---


# Effect Object (PowerPoint)

Represents timing information about a slide animation.


## Example

Use the [AddEffect](http://msdn.microsoft.com/library/fea5ac1e-83ae-2241-bf3a-8cfdd8354791%28Office.15%29.aspx)method to add an effect. This example adds a shape to the first slide in the active presentation and adds an effect and a behavior to the shape.


```
Sub NewShapeAndEffect()

    Dim shpStar As Shape

    Dim sldOne As Slide

    Dim effNew As Effect



    Set sldOne = ActivePresentation.Slides(1)

    Set shpStar = sldOne.Shapes.AddShape(Type:=msoShape5pointStar, _

        Left:=150, Top:=72, Width:=400, Height:=400)

    Set effNew = sldOne.TimeLine.MainSequence.AddEffect(Shape:=shpStar, _

        EffectId:=msoAnimEffectStretchy, Trigger:=msoAnimTriggerAfterPrevious)

    With effNew

        With .Behaviors.Add(msoAnimTypeScale).ScaleEffect

            .FromX = 75

            .FromY = 75

            .ToX = 0

            .ToY = 0

        End With

        .Timing.AutoReverse = msoTrue

    End With

End Sub
```

To refer to an existing  **Effect** object, use **[MainSequence](http://msdn.microsoft.com/library/b71f83ad-6d92-cc10-9692-a7567ca0a077%28Office.15%29.aspx)** (index), where index is the number of the **Effect** object in the **[Sequence](http://msdn.microsoft.com/library/37a5224f-2461-b575-acb6-6905bbb5136d%28Office.15%29.aspx)** collection. This example changes the effect for the first sequence and specifies the behavior for that effect.




```
Sub ChangeEffect()

    With ActivePresentation.Slides(1).TimeLine _

        .MainSequence(1)

        .EffectType = msoAnimEffectSpin

        With .Behaviors(1).RotationEffect

            .From = 100

            .To = 360

            .By = 5

        End With

    End With

End Sub
```


## Methods



|**Name**|
|:-----|
|**[Delete](http://msdn.microsoft.com/library/71261ec1-2f39-ac51-43f4-bce2b34fcadd%28Office.15%29.aspx)**|
|**[MoveAfter](http://msdn.microsoft.com/library/1d19f90c-51a6-d9bd-5593-53c67c7df415%28Office.15%29.aspx)**|
|**[MoveBefore](http://msdn.microsoft.com/library/c71f8785-737d-b2cf-8d9d-bed49e1ba754%28Office.15%29.aspx)**|
|**[MoveTo](http://msdn.microsoft.com/library/7b424225-e53c-7dc9-1e5c-14b824110027%28Office.15%29.aspx)**|

## Properties



|**Name**|
|:-----|
|**[Application](http://msdn.microsoft.com/library/031db407-eb15-2092-24b0-91bab5aab8c9%28Office.15%29.aspx)**|
|**[Behaviors](http://msdn.microsoft.com/library/e5335758-2f92-ccbc-a665-b6d5947e79f2%28Office.15%29.aspx)**|
|**[DisplayName](http://msdn.microsoft.com/library/1c8c7a78-5b09-a94e-880e-d82311cc5ee9%28Office.15%29.aspx)**|
|**[EffectInformation](http://msdn.microsoft.com/library/68c61bfc-842e-6659-eda9-cc4899c50b94%28Office.15%29.aspx)**|
|**[EffectParameters](http://msdn.microsoft.com/library/18f43203-a16e-7779-923c-7da076d2943e%28Office.15%29.aspx)**|
|**[EffectType](http://msdn.microsoft.com/library/28c2ed5f-f783-0858-cbff-8a5e6e5b8a41%28Office.15%29.aspx)**|
|**[Exit](http://msdn.microsoft.com/library/0f4d74d4-ce88-f9b9-7de5-0e42edf12967%28Office.15%29.aspx)**|
|**[Index](http://msdn.microsoft.com/library/1eac9295-e24c-c31e-3cd6-ace59f5ac04a%28Office.15%29.aspx)**|
|**[Paragraph](http://msdn.microsoft.com/library/0816387c-201d-b231-a412-ffb932c9044b%28Office.15%29.aspx)**|
|**[Parent](http://msdn.microsoft.com/library/254fa25b-ef29-c2fe-313d-daadba3e8db4%28Office.15%29.aspx)**|
|**[Shape](http://msdn.microsoft.com/library/bb392e26-1409-0a03-1cb9-c3b7c362aa7f%28Office.15%29.aspx)**|
|**[TextRangeLength](http://msdn.microsoft.com/library/b68690a5-f93e-0833-73be-a6259d604064%28Office.15%29.aspx)**|
|**[TextRangeStart](http://msdn.microsoft.com/library/b6da1565-84e2-acc4-4a06-166c5fda7071%28Office.15%29.aspx)**|
|**[Timing](http://msdn.microsoft.com/library/88b4f9a5-62aa-6844-e784-f74a1d78aa82%28Office.15%29.aspx)**|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
