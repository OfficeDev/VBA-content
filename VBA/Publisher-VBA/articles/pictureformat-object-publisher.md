---
title: PictureFormat Object (Publisher)
keywords: vbapb10.chm3670015
f1_keywords:
- vbapb10.chm3670015
ms.prod: publisher
api_name:
- Publisher.PictureFormat
ms.assetid: aa30ea9d-b91f-acdf-2e60-8a9f506f28b4
ms.date: 06/08/2017
---


# PictureFormat Object (Publisher)

Contains properties and methods that apply to pictures.


## Example

Use the  **[PictureFormat](http://msdn.microsoft.com/library/2a812ba3-18e4-fc42-6d07-535511a79650%28Office.15%29.aspx)** property to return a **PictureFormat** object. The following example sets the brightness, contrast, and color transformation for shape one on the active document and crops 18 points off the bottom of the shape. For this example to work, shape one must be either a picture or an OLE object.


```
Sub FormatPicture() 
 With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .Brightness = 0.6 
 .Contrast = 0.7 
 .ColorType = msoPictureGrayscale 
 .CropBottom = 18 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ClearCrop](http://msdn.microsoft.com/library/63a99182-d38a-6666-27ee-2641e5c43867%28Office.15%29.aspx)|
|[FillFrame](http://msdn.microsoft.com/library/404f956d-38f9-7a36-a10b-8ca8e61d59a4%28Office.15%29.aspx)|
|[FitFrame](http://msdn.microsoft.com/library/d43376ea-fd04-c8a1-011c-b2ea1be644d3%28Office.15%29.aspx)|
|[IncrementBrightness](http://msdn.microsoft.com/library/912fd08e-bbb3-bf98-b0da-7128926f3409%28Office.15%29.aspx)|
|[IncrementContrast](http://msdn.microsoft.com/library/cff50058-2b88-fc2d-633d-411380e5f2f3%28Office.15%29.aspx)|
|[Recolor](http://msdn.microsoft.com/library/42bc2280-b6d0-862a-7118-38ec1513b9c7%28Office.15%29.aspx)|
|[Remove](http://msdn.microsoft.com/library/23bfc598-166d-ee0a-aeaa-e21dd157ced3%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/b2bce79a-5c46-1473-601d-a4a25176edeb%28Office.15%29.aspx)|
|[ReplaceEx](http://msdn.microsoft.com/library/0f1b9eaf-51b6-ae21-518f-55663184ab87%28Office.15%29.aspx)|
|[RestoreOriginalColors](http://msdn.microsoft.com/library/13a0d09f-f809-a1ca-73d9-313ea293d56a%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/9ba0c997-b954-c02a-f568-c16617d5b5e5%28Office.15%29.aspx)|
|[Brightness](http://msdn.microsoft.com/library/bed1cd25-faee-6fb9-4bb3-5bdaf148b62e%28Office.15%29.aspx)|
|[ColorModel](http://msdn.microsoft.com/library/8e3e259c-943d-c1a9-f090-2ee0f0bb29f2%28Office.15%29.aspx)|
|[ColorsInPalette](http://msdn.microsoft.com/library/34e671b1-af0e-0dac-1429-246facae975b%28Office.15%29.aspx)|
|[ColorType](http://msdn.microsoft.com/library/439f9eb9-2593-d719-4ef6-0f14d1c7d0f4%28Office.15%29.aspx)|
|[Contrast](http://msdn.microsoft.com/library/f081b7c8-50cc-772b-f3b0-27c215cfebac%28Office.15%29.aspx)|
|[CropBottom](http://msdn.microsoft.com/library/8c504221-11da-f6f1-8fbb-75dc5c62b953%28Office.15%29.aspx)|
|[CropLeft](http://msdn.microsoft.com/library/f9fd2031-83f7-ea81-84eb-4f1ac6d65082%28Office.15%29.aspx)|
|[CropRight](http://msdn.microsoft.com/library/b1c20de2-e2cf-708f-ddae-194c8b1b01c1%28Office.15%29.aspx)|
|[CropTop](http://msdn.microsoft.com/library/b235898d-addf-6a4c-5693-229431545e6c%28Office.15%29.aspx)|
|[EffectiveResolution](http://msdn.microsoft.com/library/33e5323f-5e10-b2ed-62eb-03ecbbb1e893%28Office.15%29.aspx)|
|[Filename](http://msdn.microsoft.com/library/73e2a224-f15a-50cc-462e-10ccf9478122%28Office.15%29.aspx)|
|[FileSize](http://msdn.microsoft.com/library/8bad7bc0-7381-9bd8-3db8-5841e41ccb34%28Office.15%29.aspx)|
|[HasAlphaChannel](http://msdn.microsoft.com/library/97739201-cd0d-cc78-a28e-935fb11da5b3%28Office.15%29.aspx)|
|[HasTransparencyColor](http://msdn.microsoft.com/library/2e6066e8-60b0-c33e-0bb0-1b6f83208fd0%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/d98c76cc-4b75-28b7-5be1-101b372472d5%28Office.15%29.aspx)|
|[HorizontalPictureLocking](http://msdn.microsoft.com/library/9a8cb8ec-24d1-4a21-d662-bcdfd26821df%28Office.15%29.aspx)|
|[HorizontalScale](http://msdn.microsoft.com/library/7be51cde-5b2d-7870-7f39-2fa9bd714d68%28Office.15%29.aspx)|
|[ImageFormat](http://msdn.microsoft.com/library/a5523a1e-4dbf-5cd7-ba73-2a5570865ee6%28Office.15%29.aspx)|
|[IsEmpty](http://msdn.microsoft.com/library/493cbb8f-e069-14a9-a827-7f7631eb3a09%28Office.15%29.aspx)|
|[IsGreyScale](http://msdn.microsoft.com/library/1f8308c1-353e-2aac-9b4b-fad300a89b97%28Office.15%29.aspx)|
|[IsLinked](http://msdn.microsoft.com/library/2215cee8-864d-7228-8692-a428385d2be2%28Office.15%29.aspx)|
|[IsRecolored](http://msdn.microsoft.com/library/76bfbcfe-6a98-8c82-cc0a-041665aa98e6%28Office.15%29.aspx)|
|[IsTrueColor](http://msdn.microsoft.com/library/63708d40-996a-67ca-b4eb-dd53c83d1764%28Office.15%29.aspx)|
|[LeaveBlackAsBlack](http://msdn.microsoft.com/library/23b9dd90-a4aa-6659-7b08-2d1bef78e8f8%28Office.15%29.aspx)|
|[LinkedFileStatus](http://msdn.microsoft.com/library/43ddffe3-9cc3-b102-c5e8-80f26f63849c%28Office.15%29.aspx)|
|[OriginalColorsInPalette](http://msdn.microsoft.com/library/87c67430-1a5a-47f7-822f-6af8783f73b3%28Office.15%29.aspx)|
|[OriginalFileSize](http://msdn.microsoft.com/library/30704f2a-d739-7f14-d69a-73ab1f5ab8f3%28Office.15%29.aspx)|
|[OriginalHasAlphaChannel](http://msdn.microsoft.com/library/e58a97d2-4ced-d3cf-56b2-6a89df02bcdf%28Office.15%29.aspx)|
|[OriginalHeight](http://msdn.microsoft.com/library/0bf97bb1-d333-a7ed-686c-da2f3cce97c5%28Office.15%29.aspx)|
|[OriginalIsTrueColor](http://msdn.microsoft.com/library/837109d4-3479-2500-a1fa-b4c00e0f8672%28Office.15%29.aspx)|
|[OriginalResolution](http://msdn.microsoft.com/library/0cb7ee4e-3eb8-baee-6535-d936e3c5f05c%28Office.15%29.aspx)|
|[OriginalWidth](http://msdn.microsoft.com/library/3c418f3f-b2af-3176-9a37-a548b15fb4bc%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/c1d16742-a07b-04ff-4086-96da0b354f4c%28Office.15%29.aspx)|
|[RecoloredPictureColor](http://msdn.microsoft.com/library/8483c951-965d-e78d-52ff-90a16c76a5ca%28Office.15%29.aspx)|
|[TransparencyColor](http://msdn.microsoft.com/library/908d2e21-3e2a-b75b-a82d-454686b7ecb8%28Office.15%29.aspx)|
|[TransparentBackground](http://msdn.microsoft.com/library/0a78b579-92bf-36e6-22f6-3ca0a48f5b5a%28Office.15%29.aspx)|
|[VerticalPictureLocking](http://msdn.microsoft.com/library/0575d733-b515-2256-7136-6ec07532ab67%28Office.15%29.aspx)|
|[VerticalScale](http://msdn.microsoft.com/library/ff83d1bc-798b-5b42-7087-9b45f3ff573d%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/4be35ac9-a07b-b661-4be8-c4379802d711%28Office.15%29.aspx)|

