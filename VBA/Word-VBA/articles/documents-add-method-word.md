---
title: Documents.Add Method (Word)
keywords: vbawd10.chm158072846
f1_keywords:
- vbawd10.chm158072846
ms.prod: word
api_name:
- Word.Documents.Add
ms.assetid: 04b81417-cde9-4657-7737-90d266d05487
ms.date: 06/08/2017
---


# Documents.Add Method (Word)

Returns a  **Document** object that represents a new, empty document added to the collection of open documents.


## Syntax

 _expression_ . **Add**( **_Template_** , **_NewTemplate_** , **_DocumentType_** , **_Visible_** )

 _expression_ Required. A variable that represents a **[Documents](documents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Template_|Optional| **Variant**|The name of the template to be used for the new document. If this argument is omitted, the Normal template is used.|
| _NewTemplate_|Optional| **Variant**| **True** to open the document as a template. The default value is **False** .|
| _DocumentType_|Optional| **Variant**|Can be one of the following  **WdNewDocumentType** constants: **wdNewBlankDocument** , **wdNewEmailMessage** , **wdNewFrameset** , or **wdNewWebPage** . The default constant is **wdNewBlankDocument** .|
| _Visible_|Optional| **Variant**| **True** to open the document in a visible window. If this value is **False** , Microsoft Word opens the document but sets the **Visible** property of the document window to **False** . The default value is **True** .|

### Return Value

Document


## Example

This example creates a new document based on the Normal template.


```
Documents.Add
```

This example creates a new document based on the Professional Memo template.




```
Documents.Add Template:="C:\Program Files\Microsoft Office" _ 
 &; "\Templates\Memos\Professional Memo.dot"
```

This example creates and opens a new template, using the template attached to the active document as a model.




```
tmpName = ActiveDocument.AttachedTemplate.FullName 
Documents.Add Template:=tmpName, NewTemplate:=True
```


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

