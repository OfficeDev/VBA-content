---
title: ParagraphFormat Object (PowerPoint)
keywords: vbapp10.chm576000
f1_keywords:
- vbapp10.chm576000
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat
ms.assetid: 15d495cf-16e2-5cfb-e99c-a551876e3a8a
ms.date: 06/08/2017
---


# ParagraphFormat Object (PowerPoint)

Represents the paragraph formatting of a text range.


## Example

Use the [ParagraphFormat](http://msdn.microsoft.com/library/41d3f0f3-70e3-ad1a-efcb-de849d4a03d4%28Office.15%29.aspx)property to return the  **ParagraphFormat** object. The following example left aligns the paragraphs in shape two on slide one in the active presentation.


```
ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange _

    .ParagraphFormat.Alignment = ppAlignLeft
```


## Properties



|**Name**|
|:-----|
|[Alignment](http://msdn.microsoft.com/library/1083d0da-b974-f573-3306-6a865578219b%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/76be8546-409c-1762-a195-8d9c30d7a00b%28Office.15%29.aspx)|
|[BaseLineAlignment](http://msdn.microsoft.com/library/b59f680f-a5a9-f6bc-85d5-f14670269ae8%28Office.15%29.aspx)|
|[Bullet](http://msdn.microsoft.com/library/2b997a78-7791-6f08-00af-7143f94457c1%28Office.15%29.aspx)|
|[FarEastLineBreakControl](http://msdn.microsoft.com/library/ffc0cb13-b547-5a33-e661-8a2cc4237e88%28Office.15%29.aspx)|
|[HangingPunctuation](http://msdn.microsoft.com/library/e7e1f5b2-e0ed-9b5c-7c14-fcf4c134e3bb%28Office.15%29.aspx)|
|[LineRuleAfter](http://msdn.microsoft.com/library/fd206688-2217-303d-bb7e-fa3b00b0f188%28Office.15%29.aspx)|
|[LineRuleBefore](http://msdn.microsoft.com/library/2316216e-9f56-07e6-1b32-10b37a6fdc9d%28Office.15%29.aspx)|
|[LineRuleWithin](http://msdn.microsoft.com/library/0bf91b11-fe28-eec8-75f8-8fccbed19f5c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5b86ae1b-7889-0e98-43f9-7e947341edd4%28Office.15%29.aspx)|
|[SpaceAfter](http://msdn.microsoft.com/library/8b5dcf96-c35f-5e0b-6bd2-efabce7ea16f%28Office.15%29.aspx)|
|[SpaceBefore](http://msdn.microsoft.com/library/be73b3fe-4490-df58-57fd-47c51767b985%28Office.15%29.aspx)|
|[SpaceWithin](http://msdn.microsoft.com/library/523fa767-e5af-0d7f-d16a-b11dd7d3799d%28Office.15%29.aspx)|
|[TextDirection](http://msdn.microsoft.com/library/42b8cd29-c467-07c9-c9c9-f644fdc824ae%28Office.15%29.aspx)|
|[WordWrap](http://msdn.microsoft.com/library/d9ccb806-b6a0-0d4c-e272-1f15336142d1%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
