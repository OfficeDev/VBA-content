---
title: DocumentWindow Object (PowerPoint)
keywords: vbapp10.chm511000
f1_keywords:
- vbapp10.chm511000
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow
ms.assetid: 567c5e66-8d68-a868-4072-b5358cf69546
ms.date: 06/08/2017
---


# DocumentWindow Object (PowerPoint)

Represents a document window. The  **DocumentWindow** object is a member of the **[DocumentWindows](http://msdn.microsoft.com/library/84ed4b8c-593a-8100-d4b8-158115c4e84d%28Office.15%29.aspx)** collection. The **DocumentWindows** collection contains all the open document windows.


## Remarks

Use the  **[Presentation](http://msdn.microsoft.com/library/d6f5f565-d593-e230-c3b9-2302bdd83644%28Office.15%29.aspx)** property to return the presentation that's currently running in the specified document window.

Use the  **[Selection](http://msdn.microsoft.com/library/0cd670b2-53a5-87d7-8b38-761920dd9758%28Office.15%29.aspx)** property to return the selection.

Use the  **[SplitHorizontal](http://msdn.microsoft.com/library/89ec538b-d8a3-23e8-a246-35c44884a432%28Office.15%29.aspx)** property to return the percentage of the screen width that the outline pane occupies in normal view.

Use the  **[SplitVertical](http://msdn.microsoft.com/library/8a26332f-d00d-9816-30e1-48411db07a62%28Office.15%29.aspx)** property to return the percentage of the screen height that the slide pane occupies in normal view.

Use the  **[View](http://msdn.microsoft.com/library/6488ba10-744a-eb88-df8d-bf85e2f6711d%28Office.15%29.aspx)** property to return the view in the specified document window.


## Example

Use  **Windows** (index), where index is the document window index number, to return a single **DocumentWindow** object. The following example activates document window two.


```
Windows(2).Activate
```

The first member of the  **DocumentWindows** collection, `Windows(1)`, always returns the active document window. Alternatively, you can use the  **[ActiveWindow](http://msdn.microsoft.com/library/762c1c6a-1f8a-f47a-7b75-006c745caee0%28Office.15%29.aspx)** property to return the active document window. The following example maximizes the active window.




```
ActiveWindow.WindowState = ppWindowMaximized
```

Use  **Panes** (index), where index is the pane index number, to manipulate panes within normal, slide, outline, or notes page views of the document window. The following example activates pane three, which is the notes pane.




```
ActiveWindow.Panes(3).Activate
```

Use the  **[ActivePane](http://msdn.microsoft.com/library/8fa4c8a1-37b6-2676-1cfd-5fa2b130d2e3%28Office.15%29.aspx)** property to return the active pane within the document window. The following example checks to see if the active pane is the outline pane. If not, it activates the outline pane.




```
mypane = ActiveWindow.ActivePane.ViewType

    If mypane <> 1 Then

        ActiveWindow.Panes(1).Activate

    End If
```


## Methods



|**Name**|
|:-----|
|**[Activate](http://msdn.microsoft.com/library/8b6c5ede-edaf-72f2-b0f5-de2418a5e0a2%28Office.15%29.aspx)**|
|**[Close](http://msdn.microsoft.com/library/c7ba0097-5fa3-b0d0-234b-3cfe3e493522%28Office.15%29.aspx)**|
|**[ExpandSection](http://msdn.microsoft.com/library/bf4548ea-1459-9a2e-ad5a-e7d16c1b312d%28Office.15%29.aspx)**|
|**[FitToPage](http://msdn.microsoft.com/library/91ea2102-df12-20fe-cd16-e664832f9eb5%28Office.15%29.aspx)**|
|**[IsSectionExpanded](http://msdn.microsoft.com/library/ab40cd63-7daa-4406-9311-869ffd281d9a%28Office.15%29.aspx)**|
|**[LargeScroll](http://msdn.microsoft.com/library/b74ecd74-acec-0d36-68c7-1848a99fe4c1%28Office.15%29.aspx)**|
|**[NewWindow](http://msdn.microsoft.com/library/1c9f4e37-4e40-8d0b-246b-f9897ad9a56a%28Office.15%29.aspx)**|
|**[PointsToScreenPixelsX](http://msdn.microsoft.com/library/6b5f2f58-41af-3620-74f3-1c4ec3922fc2%28Office.15%29.aspx)**|
|**[PointsToScreenPixelsY](http://msdn.microsoft.com/library/0a5a96c6-3e91-31c6-ee60-ca1f8481daf0%28Office.15%29.aspx)**|
|**[RangeFromPoint](http://msdn.microsoft.com/library/74bc61e5-6c6d-0510-b549-e325dd67c7a7%28Office.15%29.aspx)**|
|**[ScrollIntoView](http://msdn.microsoft.com/library/1eee6b36-9f01-5204-dd75-1172f2e00577%28Office.15%29.aspx)**|
|**[SmallScroll](http://msdn.microsoft.com/library/f6710bca-ad85-9257-061a-dbe5829d8b7b%28Office.15%29.aspx)**|

## Properties



|**Name**|
|:-----|
|**[Active](http://msdn.microsoft.com/library/bd68b587-0811-7f40-c0da-741e2305594b%28Office.15%29.aspx)**|
|**[ActivePane](http://msdn.microsoft.com/library/8fa4c8a1-37b6-2676-1cfd-5fa2b130d2e3%28Office.15%29.aspx)**|
|**[Application](http://msdn.microsoft.com/library/89843eab-4dde-131e-85ed-a6116a98ad46%28Office.15%29.aspx)**|
|**[BlackAndWhite](http://msdn.microsoft.com/library/1363b7df-8de5-955f-60a7-682cd6b4c848%28Office.15%29.aspx)**|
|**[Caption](http://msdn.microsoft.com/library/1f0334ee-d0fa-14d4-046b-d29ffddcfd53%28Office.15%29.aspx)**|
|**[Height](http://msdn.microsoft.com/library/a81aed0f-141c-a1ca-19f0-1584680ca726%28Office.15%29.aspx)**|
|**[Left](http://msdn.microsoft.com/library/a6c8a129-b662-5fb7-4c5d-4f5d1c0aea34%28Office.15%29.aspx)**|
|**[Panes](http://msdn.microsoft.com/library/1f26709d-8414-ee89-29d8-588c6787611a%28Office.15%29.aspx)**|
|**[Parent](http://msdn.microsoft.com/library/275ed305-76f9-8dca-afb9-db206f6b128b%28Office.15%29.aspx)**|
|**[Presentation](http://msdn.microsoft.com/library/f009e2c3-aa08-09f0-c879-a25b8d1e0405%28Office.15%29.aspx)**|
|**[Selection](http://msdn.microsoft.com/library/0cd670b2-53a5-87d7-8b38-761920dd9758%28Office.15%29.aspx)**|
|**[SplitHorizontal](http://msdn.microsoft.com/library/89ec538b-d8a3-23e8-a246-35c44884a432%28Office.15%29.aspx)**|
|**[SplitVertical](http://msdn.microsoft.com/library/8a26332f-d00d-9816-30e1-48411db07a62%28Office.15%29.aspx)**|
|**[Top](http://msdn.microsoft.com/library/ba51aa9d-772a-d854-a834-60907b304e78%28Office.15%29.aspx)**|
|**[View](http://msdn.microsoft.com/library/6488ba10-744a-eb88-df8d-bf85e2f6711d%28Office.15%29.aspx)**|
|**[ViewType](http://msdn.microsoft.com/library/95eb4962-6d7a-41bd-fdae-757287f06350%28Office.15%29.aspx)**|
|**[Width](http://msdn.microsoft.com/library/ede3967a-5d52-ba5d-2279-ea7345a7d370%28Office.15%29.aspx)**|
|**[WindowState](http://msdn.microsoft.com/library/7f0ce168-0339-03f0-11e4-dc7935c04b85%28Office.15%29.aspx)**|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
