---
title: Comment Object (Excel)
keywords: vbaxl10.chm515072
f1_keywords:
- vbaxl10.chm515072
ms.prod: excel
api_name:
- Excel.Comment
ms.assetid: 3627e9be-2a28-9dc5-c822-ad42857134e3
ms.date: 06/08/2017
---


# Comment Object (Excel)

Represents a cell comment.


## Remarks

 The **Comment** object is a member of the **[Comments](comments-object-excel.md)** collection.


## Example

Use the  **[Comment](range-comment-property-excel.md)** property to return a **Comment** object. The following example changes the text in the comment in cell E5.


```
Worksheets(1).Range("E5").Comment.Text "reviewed on " &amp; Date
```

Use  **Comments** ( _index_ ), where _index_ is the comment number, to return a single comment from the **Comments** collection. The following example hides comment two on worksheet one.




```
Worksheets(1).Comments(2).Visible = False
```

Use the  **[AddComment](range-addcomment-method-excel.md)** method to add a comment to a range. The following example adds a comment to cell E5 on worksheet one.




```
With Worksheets(1).Range("e5").AddComment 
 .Visible = False 
 .Text "reviewed on " &amp; Date 
End With
```


## Methods



|**Name**|
|:-----|
|[Delete](comment-delete-method-excel.md)|
|[Next](comment-next-method-excel.md)|
|[Previous](comment-previous-method-excel.md)|
|[Text](comment-text-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](comment-application-property-excel.md)|
|[Author](comment-author-property-excel.md)|
|[Creator](comment-creator-property-excel.md)|
|[Parent](comment-parent-property-excel.md)|
|[Shape](comment-shape-property-excel.md)|
|[Visible](comment-visible-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
