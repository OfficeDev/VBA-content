---
title: Hyperlink Object (PowerPoint)
keywords: vbapp10.chm526000
f1_keywords:
- vbapp10.chm526000
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink
ms.assetid: c8d53079-b280-c93c-a3c9-b865d09abe1a
ms.date: 06/08/2017
---


# Hyperlink Object (PowerPoint)

Represents a hyperlink associated with a non-placeholder shape or text. 


## Remarks

You can use a hyperlink to jump to an Internet or intranet site, to another file, or to a slide within the active presentation. The  **Hyperlink** object is a member of the **[Hyperlinks](http://msdn.microsoft.com/library/33a3fe49-6302-0f53-22f6-b8b1594d5d57%28Office.15%29.aspx)** collection. The **Hyperlinks** collection contains all the hyperlinks on a slide or a master.


## Example

Use the [Hyperlink](http://msdn.microsoft.com/library/8654000a-bbc5-6d23-e5a7-d689bc767b1b%28Office.15%29.aspx)property to return a hyperlink for a shape. A shape can have two different hyperlinks assigned to it: one that is followed when the user clicks the shape during a slide show, and another that is followed when the user passes the mouse pointer over the shape during a slide show. For the hyperlink to be active during a slide show, the  **Action** property must be set to **ppActionHyperlink**. The following example sets the mouse-click action for shape three on slide one in the active presentation to an Internet link.


```
With ActivePresentation.Slides(1).Shapes(3) _

        .ActionSettings(ppMouseClick)

    .Action = ppActionHyperlink

    .Hyperlink.Address = "http://www.microsoft.com"

End With
```

A slide can contain more than one hyperlink. Each non-placeholder shape can have a hyperlink; the text within a shape can have its own hyperlink; and each individual character can have its own hyperlink. Use  **Hyperlinks** (index), where index is the hyperlink number, to return a single **Hyperlink** object. The following example adds the shape three mouse-click hyperlink to the Favorites folder.




```
ActivePresentation.Slides(1).Shapes(3) _

    .ActionSettings(ppMouseClick).Hyperlink.AddToFavorites
```


 **Note**  When you use this method to add a hyperlink to the Internet Explorer Favorites folder, an icon is added to the  **Favorites** menu without a corresponding name. You must add the name from within Internet Explorer.


## Methods



|**Name**|
|:-----|
|[AddToFavorites](http://msdn.microsoft.com/library/40a6f12e-3ad3-f028-ed47-b131b36af5fd%28Office.15%29.aspx)|
|[CreateNewDocument](http://msdn.microsoft.com/library/d2de9bbb-a659-3ea3-bdee-244329d88416%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/05961889-ff6c-b8f3-4cf4-e60ed782533b%28Office.15%29.aspx)|
|[Follow](http://msdn.microsoft.com/library/d56ace43-cf92-b3a6-abb4-dd7b87bc3feb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/d3d2174a-fbb2-432d-bc42-6623c91e9843%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/10191a9a-5103-f024-62dc-5fd129a56bf8%28Office.15%29.aspx)|
|[EmailSubject](http://msdn.microsoft.com/library/2416a620-9788-5da9-3095-432cab5cdc95%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5939f1a2-eb4f-d938-2683-530b0a408614%28Office.15%29.aspx)|
|[ScreenTip](http://msdn.microsoft.com/library/96ff1076-7563-8250-ea75-cee46094824e%28Office.15%29.aspx)|
|[ShowAndReturn](http://msdn.microsoft.com/library/5d08a3ff-8352-0523-2d8c-629f996b296a%28Office.15%29.aspx)|
|[SubAddress](http://msdn.microsoft.com/library/f7b34b39-6e4c-5606-8b19-92ddc0dcede5%28Office.15%29.aspx)|
|[TextToDisplay](http://msdn.microsoft.com/library/5f30033e-ddb8-8814-9e55-e0137ff6fa48%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/91a74e53-0223-ca06-6722-0bc35cda4656%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
