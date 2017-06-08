---
title: BulletFormat Object (PowerPoint)
keywords: vbapp10.chm577000
f1_keywords:
- vbapp10.chm577000
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat
ms.assetid: 8c70b2af-0175-9315-3a7e-e30aa0438798
ms.date: 06/08/2017
---


# BulletFormat Object (PowerPoint)

Represents bullet formatting.


## Example

Use the [Bullet](http://msdn.microsoft.com/library/2b997a78-7791-6f08-00af-7143f94457c1%28Office.15%29.aspx)property to return the  **BulletFormat** object. The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active presentation.


```
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

        .Visible = True

        .RelativeSize = 1.25

        .Character = 169

        With .Font

            .Color.RGB = RGB(255, 255, 0)

            .Name = "Symbol"

        End With

    End With

End With
```


## Methods



|**Name**|
|:-----|
|[Picture](http://msdn.microsoft.com/library/a38872c0-b754-bf30-3bd5-9050c5edf8f4%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/2906869e-ee3e-8a0e-9532-1bbe5cd60fef%28Office.15%29.aspx)|
|[Character](http://msdn.microsoft.com/library/42480e47-fc3a-d8aa-1368-a76b6776363a%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/4b5b6495-9e02-d8d3-c952-016561dc3f6c%28Office.15%29.aspx)|
|[Number](http://msdn.microsoft.com/library/90f92c4e-4a15-7efe-1251-5394a148db72%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/95829267-e354-828b-5034-7da64dc5d5d7%28Office.15%29.aspx)|
|[RelativeSize](http://msdn.microsoft.com/library/ce90fbcb-9aa5-a286-1f91-f06a83351b97%28Office.15%29.aspx)|
|[StartValue](http://msdn.microsoft.com/library/d243b5b4-93f6-8486-d432-a91a39ee4f81%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/2cc49660-bcf7-f965-f3cb-80e6d06ba789%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/eb1f3560-17ab-28a6-e012-cf5af292ef53%28Office.15%29.aspx)|
|[UseTextColor](http://msdn.microsoft.com/library/8242712a-051e-18fa-1b43-93a0ce1cd17b%28Office.15%29.aspx)|
|[UseTextFont](http://msdn.microsoft.com/library/8d572d8d-bd89-ec94-2484-045306d2730e%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
