---
title: WebOptions Object (Word)
keywords: vbawd10.chm2532
f1_keywords:
- vbawd10.chm2532
ms.prod: word
api_name:
- Word.WebOptions
ms.assetid: 658ae89d-3f92-067b-1309-7fc90b257111
ms.date: 06/08/2017
---


# WebOptions Object (Word)

Contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page.


## Remarks

 You can return or set attributes either at the application (global) level or at the document level. (Note that attribute values can be different from one document to another, depending on the attribute value at the time the document was saved.) Document-level attribute settings override application-level attribute settings. Application-level attributes are contained in the **DefaultWebOptions** object.

Use the  **WebOptions** property to return the **WebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and then sets the _strImageFileType_ variable accordingly.




```vb
Set objAppWebOptions = ActiveDocument.WebOptions 
With objAppWebOptions 
 If .AllowPNG = True Then 
 strImageFileType = "PNG" 
 Else 
 strImageFileType = "JPG" 
 End If 
End With
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

