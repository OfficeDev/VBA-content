---
title: Master Object (PowerPoint)
keywords: vbapp10.chm638000
f1_keywords:
- vbapp10.chm638000
ms.prod: powerpoint
api_name:
- PowerPoint.Master
ms.assetid: 22e8805e-6469-1a34-7f7b-f1ea5c6c49ff
ms.date: 06/08/2017
---


# Master Object (PowerPoint)

Represents a slide master, title master, handout master, notes master, or design master.


## Example

To return a  **Master** object, use the[Master](http://msdn.microsoft.com/library/cec5385d-f6af-dd8d-7989-251a70c4937e%28Office.15%29.aspx)property of the  **[Slide](slide-object-powerpoint.md)** object or **[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)** collection, or use the[HandoutMaster](http://msdn.microsoft.com/library/d80a8e51-61db-8da0-1fda-20a043e62569%28Office.15%29.aspx), [NotesMaster](http://msdn.microsoft.com/library/0889b69b-4c51-82cf-ccc2-ccb211d8a34e%28Office.15%29.aspx), [SlideMaster](http://msdn.microsoft.com/library/c6a9263c-462a-e9d8-7afc-32da3e133e90%28Office.15%29.aspx), or [TitleMaster](http://msdn.microsoft.com/library/d5a84b2a-fff0-dcb5-e744-466428a586b5%28Office.15%29.aspx)property of the  **[Presentation](presentation-object-powerpoint.md)** object. Note that some of these properties are also available from the **[Design](http://msdn.microsoft.com/library/3b02c779-8313-9512-c8d9-cf8a3883229f%28Office.15%29.aspx)** object as well. The following example sets the background fill for the slide master for the active presentation.


```
ActivePresentation.SlideMaster.Background.Fill _

    .PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```

To add a title master or design to a presentation and return a  **Master** object that represents the new title master or design, use the[AddTitleMaster](http://msdn.microsoft.com/library/b49baa5b-217a-ab6d-3cb3-ff74e533ef20%28Office.15%29.aspx)method. The following example adds a title master to the active presentation and places the title placeholder 10 points from the top of the master.




```
ActivePresentation.AddTitleMaster.Shapes.Title.Top = 10
```


## Methods



|**Name**|
|:-----|
|[ApplyTheme](http://msdn.microsoft.com/library/ae30318b-20e6-4eae-df4c-1f159fd77d6a%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/604d32e9-c47e-e236-de5c-7ada3e5da9ef%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ebe53ffb-cc21-fbf3-f39c-41b2d69cbf63%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/94b07efa-4e33-ac2c-c466-ff38499f81c4%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/6bd0adb1-ac97-faa0-1260-3db1bb3b3984%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/f481aa76-e96f-686a-edbb-b2bef8be0e8c%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/b42e54f7-64a5-8dcb-5079-6d6ffe8b18f0%28Office.15%29.aspx)|
|[CustomLayouts](http://msdn.microsoft.com/library/8364388f-71be-c6b7-5ab0-4150e6f62feb%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/78035fbd-e2f3-9089-2263-c04ce72394db%28Office.15%29.aspx)|
|[Guides](http://msdn.microsoft.com/library/75dabe07-406f-6770-db21-57fb8416d095%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/ac9f3282-32be-c561-e5cb-80e35db1797d%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/758cfe5a-c42c-73af-b3ed-56149275ceaa%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/5d9af48b-49e2-4253-a431-4341a697437b%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/1c751814-61fe-c246-d516-0d43b7757248%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/315325b9-c7cd-f43c-ce92-4552ff2bdd71%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/a4620f02-d3d2-da87-6bbc-430557365c2d%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/935cadd9-a57a-a792-62b4-e198546438b2%28Office.15%29.aspx)|
|[TextStyles](http://msdn.microsoft.com/library/713b6f60-5c20-6ddf-9660-4f5f2d27546d%28Office.15%29.aspx)|
|[Theme](http://msdn.microsoft.com/library/8d7852e1-2edb-9e56-2b05-f339d7436d6e%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/f57756b5-9b13-336b-0d5c-00161590ba03%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/7dd4a429-789d-fb76-2689-7e42b0668d4e%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
