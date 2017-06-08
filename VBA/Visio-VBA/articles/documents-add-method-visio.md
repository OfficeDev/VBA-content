---
title: Documents.Add Method (Visio)
keywords: vis_sdr.chm10616660
f1_keywords:
- vis_sdr.chm10616660
ms.prod: visio
api_name:
- Visio.Documents.Add
ms.assetid: 6efefc80-9373-4fe2-b290-0fff6d6bad0f
ms.date: 06/08/2017
---


# Documents.Add Method (Visio)

Adds a new  **Document** object to the **Documents** collection.


## Syntax

 _expression_ . **Add**( **_FileName_** )

 _expression_ A variable that represents a **Documents** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The type or file name of the document to add; if you do not include a path, Visio searches the folder or folders designated in the ** Application** object's **TemplatePaths** property and all published templates, including published third-party templates.|

### Return Value

Document


## Remarks

To create a new drawing based on no template, pass a zero-length string ("") to the  **Add** method.

To create a new drawing based on another file, like a template, pass the filename of the original file to the  **Add** method. Visio opens stencils that are part of the template's workspace and copies styles and other settings associated with the template to the new document. If the template file name is invalid, no document is returned and an error is generated.



To create a new stencil based on no stencil, pass ("vss").




 **Note**  Passing a filename as an argument to the  **Add** method is equivalent to opening a file like a template, where a new blank drawing is created that includes content copied from the original.


## Example

The following macro shows how to add  **Document** objects such as templates, stencils, and drawings to the **Documents** collection.

Before running this macro, replace  _Myfile.vsd_ with a valid .vsd file.




```vb
Public Sub AddDocument_Example() 
 
 Dim vsoDocument As Visio.Document 
 
 'Add a Document object based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Add a Document object based on a drawing (creates a copy of the drawing). 
 Set vsoDocument = Documents.Add("Myfile.vsd ") 
 
 'Add a Document object based on a stencil (creates a copy of the stencil). 
 Set vsoDocument = Documents.Add("Basic Shapes.vss") 
 
 'Add a Document object based on no template. 
 Set vsoDocument = Documents.Add("") 
 
End Sub
```


