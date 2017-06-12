---
title: Documents.Open Method (Visio)
keywords: vis_sdr.chm10616395
f1_keywords:
- vis_sdr.chm10616395
ms.prod: visio
api_name:
- Visio.Documents.Open
ms.assetid: 0d5e3c3f-d1bc-c7d8-0167-bb9514ede0e3
ms.date: 06/08/2017
---


# Documents.Open Method (Visio)

Opens an existing file so that it can be edited.


## Syntax

 _expression_ . **Open**( **_FileName_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of a file to open.|

### Return Value

Document


## Remarks

When you use the  **Open** method to open a **Document** object, it opens a Microsoft Visio file as an original. Depending on the file name extension, the **Open** method opens a drawing (.vsd), a stencil (.vss), a template (.vst), a workspace (.vsw), an XML drawing (.vdx), an XML stencil (.vsx), or an XML template (.vtx). You can also use this method to open and convert non-Visio files to Visio files. If the file does not exist or the file name is invalid, no **Document** object is returned and an error is generated.

If you pass a valid stencil (.vss) file name, the original stencil file opens. Starting with Microsoft Office Visio 2003, only user-created stencils are editable. By default, Visio stencils are not editable. Unless you want to create or edit the masters, open a stencil as read-only.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to open a blank document, a new document based on a template, and an existing document. 

Before running this macro, replace  _path_ \ _filename_ with the path to and file name of a valid template file (.vst) on your computer.




```vb
 
Public Sub OpenDocument_Example() 
 
 Dim vsoDocument As Visio.Document 
 
 'Open a blank document (not based on a template). 
 Set vsoDocument = Documents.Add("") 
 
 'Open a new document based on a template. 
 Set vsoDocument = Documents.Add _ 
 ("path \filename ") 
 
 'Open an existing document. 
 Set vsoDocument = Documents.Open _ 
 ("path \filename ") 
 
End Sub
```


