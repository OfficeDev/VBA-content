---
title: ParagraphFormat Object (Publisher)
keywords: vbapb10.chm5505023
f1_keywords:
- vbapb10.chm5505023
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat
ms.assetid: 0e5b1c20-564e-ef5c-f24d-1143dcaadcd8
ms.date: 06/08/2017
---


# ParagraphFormat Object (Publisher)

Represents all the formatting for a paragraph.


## Example

Use the  **[ParagraphFormat](http://msdn.microsoft.com/library/5ab0a2ec-d7a9-f3af-29e7-5421427ee783%28Office.15%29.aspx)** property to return the **ParagraphFormat** object for a paragraph or paragraphs. The **ParagraphFormat** property returns the **ParagraphFormat** object for a selection, range, or style. The following example centers the paragraph at the cursor position. This example assumes that the first shape is a text box and not another type of shape.


```
Sub CenterParagraph() 
 Selection.TextRange.ParagraphFormat _ 
 .Alignment = pbParagraphAlignmentCenter 
End Sub
```

Use the  **[Duplicate](http://msdn.microsoft.com/library/545dbfdb-4cd5-99b1-1ba3-b723e8d7b827%28Office.15%29.aspx)** property to copy an existing **ParagraphFormat** object. The following example duplicates the paragraph formatting of the first paragraph in the active publication and stores the formatting in a variable. This example duplicates an existing **ParagraphFormat** object and then changes the left indent to one inch, creates a new textbox, inserts text into it, and applies the paragraph formatting of the duplicated paragraph format to the text.




```
Sub DuplicateParagraphFormating() 
 Dim pfmtDup As ParagraphFormat 
 
 Set pfmtDup = ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.ParagraphFormat.Duplicate 
 
 pfmtDup.LeftIndent = Application.InchesToPoints(1) 
 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddTextbox(pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=200, Height:=100) 
 With .TextFrame.TextRange 
 .Text = "This is a test of how to use " &amp; _ 
 "the ParagraphFormat object." 
 .ParagraphFormat = pfmtDup 
 End With 
 End With 
 End With 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Duplicate](http://msdn.microsoft.com/library/83156999-7867-05c2-9e85-4cc0f580ac6e%28Office.15%29.aspx)|
|[Reset](http://msdn.microsoft.com/library/8ef5c799-cace-133c-33d3-3454df2c2f24%28Office.15%29.aspx)|
|[SetLineSpacing](http://msdn.microsoft.com/library/32e5b233-8415-2373-7423-18b66df3a5ea%28Office.15%29.aspx)|
|[SetListType](http://msdn.microsoft.com/library/6900aac5-fb3f-5813-309c-1422d38c8301%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Alignment](http://msdn.microsoft.com/library/db66f8b8-a813-418c-2735-e5299e6a6045%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/c8c5c15f-6cb2-86cc-a546-2616e23a1cca%28Office.15%29.aspx)|
|[AttachedToText](http://msdn.microsoft.com/library/1bfb902c-d728-1f97-513c-dcee54ce57a8%28Office.15%29.aspx)|
|[CharBasedFirstLineIndent](http://msdn.microsoft.com/library/d0432be6-2e6a-39fa-9e9a-0300a0437f35%28Office.15%29.aspx)|
|[FirstLineIndent](http://msdn.microsoft.com/library/4966b30e-7629-b66d-0870-ada91c3af4f3%28Office.15%29.aspx)|
|[KashidaPercentage](http://msdn.microsoft.com/library/d62aa512-cce6-2e78-657f-51ff1b2cbcf8%28Office.15%29.aspx)|
|[KeepLinesTogether](http://msdn.microsoft.com/library/a0f3f2f0-d986-4928-3c4f-0665711a6876%28Office.15%29.aspx)|
|[KeepWithNext](http://msdn.microsoft.com/library/fb49169d-4718-8ee6-6468-b7cbc8b8a774%28Office.15%29.aspx)|
|[LeftIndent](http://msdn.microsoft.com/library/f9cc3a86-d382-92d7-ec24-d13fc5e3d844%28Office.15%29.aspx)|
|[LineSpacing](http://msdn.microsoft.com/library/cb9abe6a-794c-6a58-2706-e12bbb5a302b%28Office.15%29.aspx)|
|[LineSpacingRule](http://msdn.microsoft.com/library/e9855daa-59f4-a4b6-f153-5de515261414%28Office.15%29.aspx)|
|[ListBulletFontName](http://msdn.microsoft.com/library/aa0269a1-c5a8-1705-551f-6b1b849701e9%28Office.15%29.aspx)|
|[ListBulletFontSize](http://msdn.microsoft.com/library/1ff1de0f-afcc-cc9c-bf45-d745695db89b%28Office.15%29.aspx)|
|[ListBulletText](http://msdn.microsoft.com/library/fa80957a-be91-398f-a24f-5a0449a9466f%28Office.15%29.aspx)|
|[ListIndent](http://msdn.microsoft.com/library/b42000ea-0636-88cf-b7ed-c71384a2b0d5%28Office.15%29.aspx)|
|[ListNumberSeparator](http://msdn.microsoft.com/library/63189011-12a0-c7bc-f6c6-7b17b0dcedf2%28Office.15%29.aspx)|
|[ListNumberStart](http://msdn.microsoft.com/library/8e17fdaa-f53e-26c4-d92b-8ead65c28555%28Office.15%29.aspx)|
|[ListType](http://msdn.microsoft.com/library/04ae7157-e864-4e95-74ff-59821eceb286%28Office.15%29.aspx)|
|[LockToBaseLine](http://msdn.microsoft.com/library/4430bab6-a338-e61d-681c-6063d4a5c3b3%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/afa92f13-222f-d48c-c739-9b21f15f1868%28Office.15%29.aspx)|
|[RightIndent](http://msdn.microsoft.com/library/bc3102d3-afc5-3f19-b98a-7f816e374d1a%28Office.15%29.aspx)|
|[SpaceAfter](http://msdn.microsoft.com/library/52f65636-862d-442e-e66f-5ff5c79ee7b0%28Office.15%29.aspx)|
|[SpaceBefore](http://msdn.microsoft.com/library/ed19a927-67e4-a1b3-06f8-1035c4b0815a%28Office.15%29.aspx)|
|[StartInNextTextBox](http://msdn.microsoft.com/library/96b34fa8-04ef-e472-16f0-15f82e7912ba%28Office.15%29.aspx)|
|[Tabs](http://msdn.microsoft.com/library/c42ba898-b84f-7215-129d-8134670f75ac%28Office.15%29.aspx)|
|[TextDirection](http://msdn.microsoft.com/library/b96c634d-0e7e-dba8-2bf4-e5baf3afa3d1%28Office.15%29.aspx)|
|[TextStyle](http://msdn.microsoft.com/library/8495c9c8-387e-a2e8-26cb-08f660dde985%28Office.15%29.aspx)|
|[UseCharBasedFirstLineIndent](http://msdn.microsoft.com/library/c2ac44ab-6671-5851-ac62-7449fd646cc5%28Office.15%29.aspx)|
|[WidowControl](http://msdn.microsoft.com/library/af1f1106-60e3-3987-3710-30fae7cf3940%28Office.15%29.aspx)|

