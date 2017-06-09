---
title: SlideShowTransition Object (PowerPoint)
keywords: vbapp10.chm539000
f1_keywords:
- vbapp10.chm539000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition
ms.assetid: 60707d0d-62a8-0366-c22f-c5c5635fd762
ms.date: 06/08/2017
---


# SlideShowTransition Object (PowerPoint)

Contains information about how the specified slide advances during a slide show.


## Example

Use the [SlideShowTransition](http://msdn.microsoft.com/library/bb931628-0ad1-e58b-9ddb-5680cb6ce9ec%28Office.15%29.aspx)property to return the  **SlideShowTransition** object. The following example specifies a Fast Strips Down-Left transition accompanied by the Bass.wav sound for slide one in the active presentation and specifies that the slide advance automatically five seconds after the previous animation or slide transition.


```
With ActivePresentation.Slides(1).SlideShowTransition

    .Speed = ppTransitionSpeedFast

    .EntryEffect = ppEffectStripsDownLeft

    .SoundEffect.ImportFromFile "c:\sndsys\bass.wav"

    .AdvanceOnTime = True

    .AdvanceTime = 5

End With

ActivePresentation.SlideShowSettings.AdvanceMode = _

    ppSlideShowUseSlideTimings
```


## Properties



|**Name**|
|:-----|
|[AdvanceOnClick](http://msdn.microsoft.com/library/0f517795-ea23-4c94-fad9-cc2e6c1cd5e6%28Office.15%29.aspx)|
|[AdvanceOnTime](http://msdn.microsoft.com/library/934c5acc-b230-2b7b-f0f2-4647cce5b62d%28Office.15%29.aspx)|
|[AdvanceTime](http://msdn.microsoft.com/library/79a120d2-5777-5eaa-a522-36e7d3bd539a%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/caf42275-9315-548a-07d9-23333ddbaaa7%28Office.15%29.aspx)|
|[Duration](http://msdn.microsoft.com/library/f8c47dda-9687-e437-8038-dae11c022914%28Office.15%29.aspx)|
|[EntryEffect](http://msdn.microsoft.com/library/4a7bb737-a977-7a02-fccf-4bbb711a6375%28Office.15%29.aspx)|
|[Hidden](http://msdn.microsoft.com/library/38e9add2-d05a-f0c3-6d8e-58e548d9789d%28Office.15%29.aspx)|
|[LoopSoundUntilNext](http://msdn.microsoft.com/library/64555d1a-20d2-cb4f-6168-dc9e9594e059%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/32ab0ea5-ad24-ba48-6c00-31a1912c8d67%28Office.15%29.aspx)|
|[SoundEffect](http://msdn.microsoft.com/library/69cff9a7-777a-57a0-d897-f132ba028bdd%28Office.15%29.aspx)|
|[Speed](http://msdn.microsoft.com/library/7c5b9dd2-88d3-5e34-619a-b35c3937a276%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
