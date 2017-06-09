---
title: Slide Object (PowerPoint)
keywords: vbapp10.chm535000
f1_keywords:
- vbapp10.chm535000
ms.prod: powerpoint
api_name:
- PowerPoint.Slide
ms.assetid: afe42344-6898-00d2-ecc1-b0ed23a71fe8
ms.date: 06/08/2017
---


# Slide Object (PowerPoint)

Represents a slide. The  **[Slides](http://msdn.microsoft.com/library/ba7f514c-8f6d-d5ef-333f-c1da0f2ab767%28Office.15%29.aspx)** collection contains all the **Slide** objects in a presentation.


## Remarks


 **Note**  Don't be confused if you're trying to return a reference to a single slide but you end up with a  **[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)** object. A single slide can be represented either by a **Slide** object or by a[SlideRange](http://msdn.microsoft.com/library/440ab59d-744a-209f-bf28-d0acd3a21e1a%28Office.15%29.aspx)collection that contains only one slide, depending on how you return a reference to the slide. For example, if you create and return a reference to a slide by using the  **[Add](http://msdn.microsoft.com/library/9a09ad9b-c52d-9fd6-20ef-68b694596ed2%28Office.15%29.aspx)** method, the slide is represented by a **Slide** object. However, if you create and return a reference to a slide by using the **[Duplicate](http://msdn.microsoft.com/library/a098ddc4-9838-35f2-86c1-8d9e4ff40209%28Office.15%29.aspx)** method, the slide is represented by a **SlideRange** collection that contains a single slide. Because all the properties and methods that apply to a **Slide** object also apply to a **SlideRange** collection that contains a single slide, you can work with the returned slide in the same way, regardless of whether it is represented by a **Slide** object or a **SlideRange** collection.

The following examples describe how to:


- Return a slide that you specify by name, index number, or slide ID number
    
- Return a slide in the selection
    
- Return the slide that's currently displayed in any document window or slide show window you specify
    
- Create a new slide
    

## Example

Use  **Slides** (index), where index is the slide name or index number, or use **Slides.FindBySlideID** (index), where index is the slide ID number, to return a single **Slide** object. The following example sets the layout for slide one in the active presentation.


```
ActivePresentation.Slides(1).Layout = ppLayoutTitle
```

The following example sets the layout for the slide with the ID number 265.




```
ActivePresentation.Slides.FindBySlideID(265).Layout = ppLayoutTitle
```

Use  **Selection.SlideRange** (index), where index is the slide name or index number within the selection, to return a single **Slide** object. The following example sets the layout for slide one in the selection in the active window, assuming that there's at least one slide selected.




```
ActiveWindow.Selection.SlideRange(1).Layout = ppLayoutTitle
```

If there's only one slide selected, you can use  **Selection.SlideRange** to return a **SlideRange** collection that contains the selected slide. The following example sets the layout for slide one in the current selection in the active window, assuming that there's exactly one slide selected.




```
ActiveWindow.Selection.SlideRange.Layout = ppLayoutTitle
```

Use the  **Slide** property to return the slide that's currently displayed in the specified document window or slide show window view. The following example copies the slide that's currently displayed in document window two to the Clipboard.




```
Windows(2).View.Slide.Copy
```

Use the  **Add** method to create a new slide and add it to the presentation. The following example adds a title slide to the beginning of the active presentation.




```
ActivePresentation.Slides.Add 1, ppLayoutTitleOnly
```


## Methods



|**Name**|
|:-----|
|[ApplyTemplate](http://msdn.microsoft.com/library/ecefec47-697e-57d6-375c-47ccd80268a4%28Office.15%29.aspx)|
|[ApplyTemplate2](http://msdn.microsoft.com/library/e4931f7b-98de-a854-3752-c1f9ca70cf3b%28Office.15%29.aspx)|
|[ApplyTheme](http://msdn.microsoft.com/library/70fff6cd-0541-dff8-754e-e8ee1a46dc2b%28Office.15%29.aspx)|
|[ApplyThemeColorScheme](http://msdn.microsoft.com/library/30a29534-d2ea-0f7e-8905-85c82ab4c1a9%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/35844287-a2f3-463d-f735-d88f383ad208%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/03029017-52c8-5176-a218-8b5ff8edec10%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/1b59cab0-cd3c-6d86-5207-a637557e3fcc%28Office.15%29.aspx)|
|[Duplicate](http://msdn.microsoft.com/library/a098ddc4-9838-35f2-86c1-8d9e4ff40209%28Office.15%29.aspx)|
|[Export](http://msdn.microsoft.com/library/b7379dfa-ce0b-340d-9109-5970beb77aa3%28Office.15%29.aspx)|
|[MoveTo](http://msdn.microsoft.com/library/b044a6fe-b6af-0f7f-ca4a-69d8a6f146e6%28Office.15%29.aspx)|
|[MoveToSectionStart](http://msdn.microsoft.com/library/757a0e42-85d1-2b03-65f7-92d15c626320%28Office.15%29.aspx)|
|[PublishSlides](http://msdn.microsoft.com/library/76f7bd2a-f48c-33e5-52dc-ae9757a880db%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/8c9511bd-4d21-fe81-f2b9-38ffef028d63%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ef89143b-2a7e-b7b3-a790-3bcb7433c1fd%28Office.15%29.aspx)|
|[Background](http://msdn.microsoft.com/library/8af622b9-029a-6839-7a44-fdf96fe75dc9%28Office.15%29.aspx)|
|[BackgroundStyle](http://msdn.microsoft.com/library/5f085f74-8f67-94fa-213e-46be866155fe%28Office.15%29.aspx)|
|[ColorScheme](http://msdn.microsoft.com/library/3d40d93f-4e7d-e95f-8340-d138da2a1b55%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/396c2d6b-f0cb-3ed8-94ae-6ee864d194c1%28Office.15%29.aspx)|
|[CustomerData](http://msdn.microsoft.com/library/4a31363b-9fcb-e062-3bf1-f31090ee2d29%28Office.15%29.aspx)|
|[CustomLayout](http://msdn.microsoft.com/library/0dcf50e8-b09a-c1da-4e72-50797eb09f9c%28Office.15%29.aspx)|
|[Design](http://msdn.microsoft.com/library/bac64534-92f7-5611-db7e-501504e577e1%28Office.15%29.aspx)|
|[DisplayMasterShapes](http://msdn.microsoft.com/library/9a4a5146-e84d-b9fe-a837-0bcafa3fe61d%28Office.15%29.aspx)|
|[FollowMasterBackground](http://msdn.microsoft.com/library/252c1893-f877-082a-8778-4ee9cc1d9c72%28Office.15%29.aspx)|
|[HasNotesPage](http://msdn.microsoft.com/library/5c92e382-ffe0-c4c4-7989-5ac84e82adc0%28Office.15%29.aspx)|
|[HeadersFooters](http://msdn.microsoft.com/library/947eb2cf-6902-2eb1-f781-0602e96bbdef%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/0e1d7545-815f-3be9-38b8-355f9e6e9962%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/681819b8-327e-fb6f-e9d2-0f8feb48ec36%28Office.15%29.aspx)|
|[Master](http://msdn.microsoft.com/library/cec5385d-f6af-dd8d-7989-251a70c4937e%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/11d6a295-02b6-3cf2-0e8b-42637e3b1f11%28Office.15%29.aspx)|
|[NotesPage](http://msdn.microsoft.com/library/8d102704-1660-cc5f-6701-d7bc67b5924b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/02925312-0c0b-b1b9-c353-7d559f0e0050%28Office.15%29.aspx)|
|[PrintSteps](http://msdn.microsoft.com/library/b5474b85-0c1f-aa18-da9d-be7d778e9e16%28Office.15%29.aspx)|
|[sectionIndex](http://msdn.microsoft.com/library/4a992a39-100a-d23b-0a67-c24199ff9a9f%28Office.15%29.aspx)|
|[Shapes](http://msdn.microsoft.com/library/8eaf3611-2799-835d-ecaa-c8f802256673%28Office.15%29.aspx)|
|[SlideID](http://msdn.microsoft.com/library/9d2d920c-a876-c71c-083f-ae8a3ad06c85%28Office.15%29.aspx)|
|[SlideIndex](http://msdn.microsoft.com/library/8a046547-9655-7281-a406-1533f41016aa%28Office.15%29.aspx)|
|[SlideNumber](http://msdn.microsoft.com/library/6d62848b-5969-c711-9df4-2b9140ec502c%28Office.15%29.aspx)|
|[SlideShowTransition](http://msdn.microsoft.com/library/bb931628-0ad1-e58b-9ddb-5680cb6ce9ec%28Office.15%29.aspx)|
|[Tags](http://msdn.microsoft.com/library/2869e5db-3355-0747-633b-2da430667e5b%28Office.15%29.aspx)|
|[ThemeColorScheme](http://msdn.microsoft.com/library/aaa8f7b5-e7c9-6c75-d88b-858a5dd3429d%28Office.15%29.aspx)|
|[TimeLine](http://msdn.microsoft.com/library/7dda6e00-5e22-fb2f-91d9-e9c15f8d62bd%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
