---
title: XMLSchemaReferences Object (Word)
keywords: vbawd10.chm1772
f1_keywords:
- vbawd10.chm1772
ms.prod: word
api_name:
- Word.XMLSchemaReferences
ms.assetid: 56bef973-805c-c77a-6d2a-54a39fbd1206
ms.date: 06/08/2017
---


# XMLSchemaReferences Object (Word)

A collection of  **XMLSchemaReference** objects that represent the unique namespaces that are attached to a document.


## Remarks

Use the  **XMLSchemaReferences** property to return a collection of schemas attached to a document. The following example loops through the schemas attached to a document. If it finds the specified schema, it reloads it; if it doesn't find the specified schema, it attaches the schema to the document.


```vb
Sub VerifySampleSchema() 
 Dim objNS As XMLNamespace 
 Dim objSchema As XMLSchemaReference 
 Dim blnSchemaAttached As Boolean 
 
 For Each objSchema In ActiveDocument.XMLSchemaReferences 
 If objSchema.NamespaceURI <> "SimpleSample" Then 
 blnSchemaAttached = False 
 Else 
 objSchema.Reload 
 blnSchemaAttached = True 
 Exit For 
 End If 
 Next 
 
 If blnSchemaAttached = False Then 
 Set objNS = Application.XMLNamespaces.Item("SimpleSample") 
 objNS.AttachToDocument (ActiveDocument) 
 End If 
End Sub
```


 **Note**  The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK on the Microsoft Developer Network (MSDN) Web site.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


