---
title: Hyperlink Object (Publisher)
keywords: vbapb10.chm4653055
f1_keywords:
- vbapb10.chm4653055
ms.prod: publisher
api_name:
- Publisher.Hyperlink
ms.assetid: 1cc6d95b-357a-c169-a5d2-6850a1a3bbd6
ms.date: 06/08/2017
---


# Hyperlink Object (Publisher)

Represents a hyperlink. The  **Hyperlink** object is a member of the **[Hyperlinks](hyperlinks-object-publisher.md)** collection and the **[Shape](http://msdn.microsoft.com/library/666cb7f0-62a8-f419-9838-007ef29506ee%28Office.15%29.aspx)** and **[ShapeRange](shaperange-object-publisher.md)** objects.


## Example

Use the  **[Hyperlink](http://msdn.microsoft.com/library/0990ab32-b4a3-6c89-cb9f-8f8c64ef804f%28Office.15%29.aspx)** property to return a **Hyperlink** object associated with a shape (a shape can have only one hyperlink). The following example deletes the hyperlink associated with the first shape in the active document.


```
Sub DeleteHyperlink() 
 ActiveDocument.Pages(1).Shapes(1).Hyperlink.Delete 
End Sub
```

Use  **Hyperlinks** (index), where index is the index number, to return a single **Hyperlink** object from a document, range, or selection. The following example deletes the first hyperlink in the selection.




```
Sub DeleteSelectedHyperlink() 
 If Selection.TextRange.Hyperlinks.Count >= 1 Then 
 Selection.TextRange.Hyperlinks(1).Delete 
 End If 
End Sub
```

Use the  **[Add](http://msdn.microsoft.com/library/f5a8cc01-a571-623d-bfab-fe48e43a21b1%28Office.15%29.aspx)** method to add a hyperlink. The following example adds a hyperlink to the selected text.




```
Sub AddHyperlinkToSelectedText() 
 Selection.TextRange.Hyperlinks.Add Text:=Selection.TextRange, _ 
 Address:="http://www.tailspintoys.com/" 
End Sub
```

Use the  **[Address](http://msdn.microsoft.com/library/784a9213-38bc-c5fd-f215-abeb174ec628%28Office.15%29.aspx)** property to add or change the address to a hyperlink. The following example adds a shape to the active publication and then adds a hyperlink to the shape.




```
Sub AddHyperlinkToShape() 
 With ActiveDocument.Pages(1).Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=200, _ 
 Top:=200, Width:=300, Height:=300) 
 .Hyperlink.Address = "http://www.tailspintoys.com/" 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/28b7f351-c1a8-29f1-2114-ed6854fbd13a%28Office.15%29.aspx)|
|[SetPageRelative](http://msdn.microsoft.com/library/4b2f2e84-09ce-cef6-6f22-b82642cc71fe%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/784a9213-38bc-c5fd-f215-abeb174ec628%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/dadf9b35-580e-c184-c439-38b3a4f1529f%28Office.15%29.aspx)|
|[EmailSubject](http://msdn.microsoft.com/library/16b60648-56fe-b8ba-3424-0dd6e88727e6%28Office.15%29.aspx)|
|[PageID](http://msdn.microsoft.com/library/1b5051eb-e6b4-a5a7-610a-5be03863a92b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/a0e3ab66-cdc4-09ab-6995-8a5e0194d6e2%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/ff105ffe-cb48-0f6a-99ff-eaac0500938f%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/afd1dab7-472a-2aa5-f5da-1e2f783b5270%28Office.15%29.aspx)|
|[TargetType](http://msdn.microsoft.com/library/1cbc8c36-563c-4464-4f0d-2836682ce532%28Office.15%29.aspx)|
|[TextToDisplay](http://msdn.microsoft.com/library/26b5857c-3f94-0d33-f65e-9c34f2a4cc2b%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/6a9ac3c4-4f34-d759-af95-a3bdc510a56f%28Office.15%29.aspx)|

