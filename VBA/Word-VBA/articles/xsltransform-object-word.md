---
title: XSLTransform Object (Word)
keywords: vbawd10.chm1171
f1_keywords:
- vbawd10.chm1171
ms.prod: word
api_name:
- Word.XSLTransform
ms.assetid: cccf0383-8b21-0f46-b5b6-9a092599fd76
ms.date: 06/08/2017
---


# XSLTransform Object (Word)

Represents a single registered Extensible Stylesheet Language Transformation (XSLT).


## Remarks

Use the  **Add** method of the **XSLTransforms** collection to add an individual XSLT to the list of XSLTs available for a schema. The following example adds the simplesample.xslt transformation to the XSLTs for the SimpleSample schema.


```vb
Sub AddXSLT() 
 Dim objSchema As XMLNamespace 
 Dim objTransform As XSLTransform 
 
 Set objSchema = Application.XMLNamespaces("SimpleSample") 
 Set objTransform = objSchema.XSLTransforms _ 
 .Add("c:\schemas\simplesample.xslt") 
 
End Sub
```

Use the  **Item** method of the **XSLTransforms** collection to return a single **XSLTransform** object. The following example deletes the first XSLT in the collection of XSLTs for the SimpleSample schema.




```vb
Sub DeleteTransform() 
 Dim objXSLT As XSLTransform 
 Dim intResponse As Integer 
 
 Set objXSLT = Application.XMLNamespaces("SimpleSample") _ 
 .XSLTransforms.Item(1) 
 
 intResponse = MsgBox("Are you sure you want to delete the " _ 
 &; objXSLT.Alias &; " XSLT?", vbYesNo) 
 
 If intResponse = vbYes Then objXSLT.Delete 
 
End Sub
```


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


