---
title: Comments Object (Excel)
keywords: vbaxl10.chm513072
f1_keywords:
- vbaxl10.chm513072
ms.prod: excel
api_name:
- Excel.Comments
ms.assetid: f43bf021-1e46-10cf-09bf-070fc6a2c81a
ms.date: 06/08/2017
---


# Comments Object (Excel)

A collection of cell comments.


## Remarks

 Each comment is represented by a **[Comment](comment-object-excel.md)** object.


## Example

Use the  **Comments** property to return the **Comments** collection. The following example hides all the comments on worksheet one.


```vb
Set cmt = Worksheets(1).Comments 
For Each c In cmt 
 c.Visible = False 
Next
```

Use the  **[AddComment](range-addcomment-method-excel.md)** method to add a comment to a range. The following example adds a comment to cell E5 on worksheet one.




```vb
With Worksheets(1).Range("e5").AddComment 
 .Visible = False 
 .Text "reviewed on " &; Date 
End With
```

Use  **Comments** ( _index_ ), where _index_ is the comment number, to return a single comment from the **Comments** collection. The following example hides comment two on worksheet one.




```vb
Worksheets(1).Comments(2).Visible = False
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

