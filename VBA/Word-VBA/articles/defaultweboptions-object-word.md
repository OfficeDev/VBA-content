---
title: DefaultWebOptions Object (Word)
ms.prod: word
api_name:
- Word.DefaultWebOptions
ms.assetid: 7459af1e-c495-f84f-929c-f7a611ec49b3
ms.date: 06/08/2017
---


# DefaultWebOptions Object (Word)

Contains global application-level attributes used by Microsoft Word when you open a Web page or save a document as a Web page.


## Remarks

You can return or set attributes either at the application (global) level or at the document level. (Note that attribute values can be different from one document to another, depending on the attribute value at the time the document was saved.) Document-level attribute settings override application-level attribute settings. Document-level attributes are contained in the  **[WebOptions](weboptions-object-word.md)** object.

Use the  **DefaultWebOptions** method to return the **DefaultWebOptions** object. The following example checks to see whether PNG (Portable Network Graphics) is allowed as an image format and sets the _strImageFileType_ variable accordingly.




```vb
Set objAppWebOptions = Application.DefaultWebOptions 
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


