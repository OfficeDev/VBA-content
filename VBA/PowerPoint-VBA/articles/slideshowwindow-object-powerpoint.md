---
title: SlideShowWindow Object (PowerPoint)
keywords: vbapp10.chm507000
f1_keywords:
- vbapp10.chm507000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindow
ms.assetid: 22468489-d4a2-ffea-7479-53ecb8d5da29
ms.date: 06/08/2017
---


# SlideShowWindow Object (PowerPoint)

Represents a window in which a slide show runs.


## Example

Use  **SlideShowWindows** (index), where index is the slide show window index number, to return a single **SlideShowWindow** object. The following example activates slide show window two.


```
SlideShowWindows(2).Activate
```

Use the [Run](http://msdn.microsoft.com/library/497fae3b-b6a3-dc26-20d9-bdc8057ddc09%28Office.15%29.aspx)method to create a new slide show window and return a reference to this slide show window. The following example runs a slide show of the active presentation and reduces the height of the slide show window just enough so that you can see the taskbar (for monitors with a screen resolution of 800 by 600).




```
With ActivePresentation.SlideShowSettings

    .ShowType = ppShowTypeSpeaker

    With .Run

        .Height = 300

        .Width = 400

    End With

End With
```

Use the [View](http://msdn.microsoft.com/library/ebf565af-fc90-ab1b-0e05-6dcb90a7c2d2%28Office.15%29.aspx)property to return the view in the specified slide show window. The following example sets the view in slide show window one to display slide three in the presentation.




```
SlideShowWindows(1).View.GotoSlide 3
```

Use the [Presentation](http://msdn.microsoft.com/library/9c05deb7-a385-540f-97a5-1c5510f120c6%28Office.15%29.aspx)property to return the presentation that's currently running in the specified slide show window. The following example displays the name of the presentation that's currently running in slide show window one.




```
MsgBox SlideShowWindows(1).Presentation.Name
```


## Methods



|**Name**|
|:-----|
|[DrawLine](http://msdn.microsoft.com/library/d4c3c1c9-cd12-67ba-b1b9-4d7e924bd084%28Office.15%29.aspx)|
|[EndNamedShow](http://msdn.microsoft.com/library/1b829558-a729-8aa1-c260-8b7410501153%28Office.15%29.aspx)|
|[EraseDrawing](http://msdn.microsoft.com/library/d1ccb77b-c591-f3ec-bb88-1f317f057103%28Office.15%29.aspx)|
|[Exit](http://msdn.microsoft.com/library/9abcb628-395b-02bf-3a61-d0c7b8429741%28Office.15%29.aspx)|
|[First](http://msdn.microsoft.com/library/5f360832-2deb-b3df-7b55-5a3c964d0057%28Office.15%29.aspx)|
|[FirstAnimationIsAutomatic](http://msdn.microsoft.com/library/689b2dfc-a441-51c6-9eea-de99194ba203%28Office.15%29.aspx)|
|[GetClickCount](http://msdn.microsoft.com/library/3df28d31-4da1-1ea3-e1d6-5ff334018ebc%28Office.15%29.aspx)|
|[GetClickIndex](http://msdn.microsoft.com/library/678feca3-79d4-e4e8-83aa-3484f5c099e9%28Office.15%29.aspx)|
|[GotoClick](http://msdn.microsoft.com/library/b41dec86-96a9-447a-5895-0b28fc4bd6b2%28Office.15%29.aspx)|
|[GotoNamedShow](http://msdn.microsoft.com/library/7e26b77f-bb7b-fd32-eabf-bc8f568e5c62%28Office.15%29.aspx)|
|[GotoSlide](http://msdn.microsoft.com/library/f733f46d-a632-02cb-3dbf-f29122fe347a%28Office.15%29.aspx)|
|[Last](http://msdn.microsoft.com/library/1188d75f-9561-b92c-e2d1-9ceb03eae904%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/cf95eef7-4fd7-4c47-4436-037ec1882d4c%28Office.15%29.aspx)|
|[Player](http://msdn.microsoft.com/library/d7bb6b02-516b-07bb-42b4-ae245ce20262%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/a53741b0-8325-696c-51e5-ffd3f9358ca8%28Office.15%29.aspx)|
|[ResetSlideTime](http://msdn.microsoft.com/library/aa00c585-d3c3-9cdc-860d-8c1f2f0a6ef3%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AcceleratorsEnabled](http://msdn.microsoft.com/library/04db702f-af30-1868-0cab-17e692892e82%28Office.15%29.aspx)|
|[AdvanceMode](http://msdn.microsoft.com/library/cdc2a780-c591-b96d-cc2e-7b0571056491%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/bdfbaf89-cd91-2a3a-481c-346c11b889e7%28Office.15%29.aspx)|
|[CurrentShowPosition](http://msdn.microsoft.com/library/390eb2c3-059f-f7e9-e91a-0e8cf9a0ddff%28Office.15%29.aspx)|
|[IsNamedShow](http://msdn.microsoft.com/library/a68632b2-bff4-9047-f0b8-6acb22a29071%28Office.15%29.aspx)|
|[LaserPointerEnabled](http://msdn.microsoft.com/library/9ba56542-a2bf-28d2-9609-50f9a4144c91%28Office.15%29.aspx)|
|[LastSlideViewed](http://msdn.microsoft.com/library/47647e03-d898-47b5-cb50-79f3e368b56f%28Office.15%29.aspx)|
|[MediaControlsHeight](http://msdn.microsoft.com/library/523732d6-6b6a-7658-a8f0-dbdeb9e3e68e%28Office.15%29.aspx)|
|[MediaControlsLeft](http://msdn.microsoft.com/library/1cc3c3a2-63d8-e43b-2056-3638caa039fe%28Office.15%29.aspx)|
|[MediaControlsTop](http://msdn.microsoft.com/library/e530dad8-ab23-e37d-fde3-5edb79c51365%28Office.15%29.aspx)|
|[MediaControlsVisible](http://msdn.microsoft.com/library/0d9d9807-bd5f-4633-001f-9aa4f63c5c28%28Office.15%29.aspx)|
|[MediaControlsWidth](http://msdn.microsoft.com/library/02a81c3e-c19d-183a-c9e4-08decf01d30f%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0e21d9e5-48d3-2a4c-fe64-8a33e4341417%28Office.15%29.aspx)|
|[PointerColor](http://msdn.microsoft.com/library/29f4c5e0-0927-1dbb-7bc9-b147ae38ff88%28Office.15%29.aspx)|
|[PointerType](http://msdn.microsoft.com/library/58f40da1-ae25-4604-86bc-6fb884b8fd16%28Office.15%29.aspx)|
|[PresentationElapsedTime](http://msdn.microsoft.com/library/6f710354-1691-4673-f83f-395d510d6999%28Office.15%29.aspx)|
|[Slide](http://msdn.microsoft.com/library/4fdee96b-9b0d-64ba-19de-b810bf07987b%28Office.15%29.aspx)|
|[SlideElapsedTime](http://msdn.microsoft.com/library/e9250ea3-c37e-ebed-c8a8-9774dab77f37%28Office.15%29.aspx)|
|[SlideShowName](http://msdn.microsoft.com/library/63efa2d8-7321-dc72-3c25-ab5ab4ba5c0a%28Office.15%29.aspx)|
|[State](http://msdn.microsoft.com/library/749fe106-fed4-6ccc-f127-2e8a80196309%28Office.15%29.aspx)|
|[Zoom](http://msdn.microsoft.com/library/92a303f0-b37f-a017-bedb-6537e235f753%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
