---
title: Series Object (PowerPoint)
keywords: vbapp10.chm716000
f1_keywords:
- vbapp10.chm716000
ms.prod: powerpoint
api_name:
- PowerPoint.Series
ms.assetid: 5c8c2d92-d8ca-4d21-e213-c374292275d4
ms.date: 06/08/2017
---


# Series Object (PowerPoint)

Represents a series in a chart.


## Remarks

 The **Series** object is a member of the **[SeriesCollection](seriescollection-object-powerpoint.md)** collection.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **[SeriesCollection](http://msdn.microsoft.com/library/8adeb8b4-ba4f-6cdf-33bf-dceb1845dfb8%28Office.15%29.aspx)** ( _Index_ ), where _Index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series of the first chart in the active document.

The series index number indicates the order in which the series were added to the chart.  `SeriesCollection(1)` is the first series added to the chart, and `SeriesCollection(SeriesCollection.Count)` is the last one added.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[ApplyDataLabels](http://msdn.microsoft.com/library/d8f4752f-1ff4-8a42-4b9f-12d81814f4f2%28Office.15%29.aspx)|
|[ClearFormats](http://msdn.microsoft.com/library/068e8908-9e88-52e9-0e44-1260b7ad21c6%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/7725e3f1-a3a8-9d03-db25-ef6b6ef31caf%28Office.15%29.aspx)|
|[DataLabels](http://msdn.microsoft.com/library/e1e37006-8a4d-9a55-02a4-890ec5e608db%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/36684621-b198-689a-d7b2-9dbaf2a7f8c3%28Office.15%29.aspx)|
|[ErrorBar](http://msdn.microsoft.com/library/a25795b8-a954-0803-bea6-6c650190ad3f%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/3f74aabb-f9c0-c76d-eaaa-c08c21daef48%28Office.15%29.aspx)|
|[Points](http://msdn.microsoft.com/library/53bec845-d3a0-fdce-921b-66d2d4e1eb59%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/13b8b940-c05c-bcaa-8cba-5a63e2445d51%28Office.15%29.aspx)|
|[Trendlines](http://msdn.microsoft.com/library/17578607-d0aa-dcc2-1eec-3af031f17c2d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/e6a8c8a0-928a-08b5-82e0-1ea060cb0daf%28Office.15%29.aspx)|
|[ApplyPictToEnd](http://msdn.microsoft.com/library/fa71354c-c76a-545a-ae3c-22ae36260365%28Office.15%29.aspx)|
|[ApplyPictToFront](http://msdn.microsoft.com/library/babe864c-1301-a8d1-ab13-41b9ccc71824%28Office.15%29.aspx)|
|[ApplyPictToSides](http://msdn.microsoft.com/library/b8a5b93d-f674-3927-3742-7578656f3152%28Office.15%29.aspx)|
|[AxisGroup](http://msdn.microsoft.com/library/c08c5039-eea6-5fed-a1b8-8c18b4886439%28Office.15%29.aspx)|
|[BarShape](http://msdn.microsoft.com/library/c6f2d0b7-407e-4073-89b1-433e9386287a%28Office.15%29.aspx)|
|[BubbleSizes](http://msdn.microsoft.com/library/c4be04b4-fb9c-1301-a5cb-e16528a97903%28Office.15%29.aspx)|
|[ChartType](http://msdn.microsoft.com/library/2ee70821-c909-bd90-a07f-7520be7b3117%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/36a05700-cbf8-0114-633d-70cf6991514a%28Office.15%29.aspx)|
|[ErrorBars](http://msdn.microsoft.com/library/6d3a4bd3-93f1-95d6-6d8e-4f296c1b5f95%28Office.15%29.aspx)|
|[Explosion](http://msdn.microsoft.com/library/c952b296-35c2-0215-228e-883a29e1b9d8%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/2c1e7a2e-6f2e-7b18-c29b-cec3ba61f563%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/04d62f5d-e63d-1643-a6cd-eae0c37b73cf%28Office.15%29.aspx)|
|[FormulaLocal](http://msdn.microsoft.com/library/93f20166-0d98-a05e-6938-dfc18f46e936%28Office.15%29.aspx)|
|[FormulaR1C1](http://msdn.microsoft.com/library/26b5e5e1-bcc2-a9f6-1767-dec6959901a9%28Office.15%29.aspx)|
|[FormulaR1C1Local](http://msdn.microsoft.com/library/cb00cca5-b540-6083-7fc5-2d2d6a58719f%28Office.15%29.aspx)|
|[Has3DEffect](http://msdn.microsoft.com/library/ce72d83a-d89e-1953-980e-3caea6b4d4c9%28Office.15%29.aspx)|
|[HasDataLabels](http://msdn.microsoft.com/library/b0b9bd37-7416-9903-d656-c4e468a9e481%28Office.15%29.aspx)|
|[HasErrorBars](http://msdn.microsoft.com/library/658e45b6-0c1c-af50-491a-d88468782227%28Office.15%29.aspx)|
|[HasLeaderLines](http://msdn.microsoft.com/library/4aaab32e-56e7-cd47-c3a2-ff92df218373%28Office.15%29.aspx)|
|[InvertColor](http://msdn.microsoft.com/library/e2ca8473-11d0-98fe-587e-740f7a00e85b%28Office.15%29.aspx)|
|[InvertColorIndex](http://msdn.microsoft.com/library/879637a8-52a7-a6ac-a882-386dad1808cb%28Office.15%29.aspx)|
|[InvertIfNegative](http://msdn.microsoft.com/library/dd672a13-d419-c68f-3330-a1449d14f636%28Office.15%29.aspx)|
|[IsFiltered](http://msdn.microsoft.com/library/1a349eac-0fa0-3bdb-cdf4-ab25b8e37189%28Office.15%29.aspx)|
|[LeaderLines](http://msdn.microsoft.com/library/f5c706e0-c6df-ae45-9f34-b7f6b4200326%28Office.15%29.aspx)|
|[MarkerBackgroundColor](http://msdn.microsoft.com/library/6cd480e7-c291-7c11-1d3f-57099805d2c0%28Office.15%29.aspx)|
|[MarkerBackgroundColorIndex](http://msdn.microsoft.com/library/18640945-ac4a-c661-46fa-804a66f57502%28Office.15%29.aspx)|
|[MarkerForegroundColor](http://msdn.microsoft.com/library/3d312b67-7fcf-5446-c57d-9831af908e8d%28Office.15%29.aspx)|
|[MarkerForegroundColorIndex](http://msdn.microsoft.com/library/85535a03-fb8c-fe76-9b67-ef60d51987b1%28Office.15%29.aspx)|
|[MarkerSize](http://msdn.microsoft.com/library/60a402b8-69f5-db47-73df-55ed75a42272%28Office.15%29.aspx)|
|[MarkerStyle](http://msdn.microsoft.com/library/e985978e-f0cf-b809-ebe1-f5504e9e8df6%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/848bdef3-76fc-2993-bbc3-4925bccbb1b9%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/1549d1bb-0a61-7772-fae4-3e6c941b7276%28Office.15%29.aspx)|
|**[ParentDataLabelOption](http://msdn.microsoft.com/library/678ad97d-725b-5a4c-b3a4-294e9f905e5f%28Office.15%29.aspx)**|
|:-----|
|[PictureType](http://msdn.microsoft.com/library/106933a2-49a7-e9d3-e5fa-fd2d0ab8974a%28Office.15%29.aspx)|
|[PictureUnit2](http://msdn.microsoft.com/library/83ccb10a-1883-9665-8a63-4494e853aa72%28Office.15%29.aspx)|
|[PlotColorIndex](http://msdn.microsoft.com/library/84d9a44b-7841-ca68-74e8-62537e534ed8%28Office.15%29.aspx)|
|[PlotOrder](http://msdn.microsoft.com/library/196c0b37-a9fe-ec01-ca0a-786c70e8a63c%28Office.15%29.aspx)|
|[QuartileCalculationInclusiveMedian](http://msdn.microsoft.com/library/0c6e80be-22f6-8e7e-437c-7c9066e0886d%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/4b530abf-5966-89eb-3ab2-5dbe4c1f2adf%28Office.15%29.aspx)|
|[Smooth](http://msdn.microsoft.com/library/fff72f72-25f3-801c-67eb-b801102c8aed%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/87dcb817-fd6d-d249-cd8d-50cbfe051cf0%28Office.15%29.aspx)|
|[Values](http://msdn.microsoft.com/library/ff6ceb5c-e7c3-6b75-8225-d18dd3baa2b8%28Office.15%29.aspx)|
|[XValues](http://msdn.microsoft.com/library/e1e83dc0-ed73-c29b-942a-575511ce94e1%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
