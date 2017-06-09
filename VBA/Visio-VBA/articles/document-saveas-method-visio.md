---
title: Document.SaveAs Method (Visio)
keywords: vis_sdr.chm10516500
f1_keywords:
- vis_sdr.chm10516500
ms.prod: visio
api_name:
- Visio.Document.SaveAs
ms.assetid: 308e92b1-de61-9ce3-19be-b7f9126247a0
ms.date: 06/08/2017
---


# Document.SaveAs Method (Visio)

Saves a document and gives it a file name.


## Syntax

 _expression_ . **SaveAs**( **_FileName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name for the document.|

### Return Value

Integer


## Remarks

The  **SaveAs** method can accept drive names that use the universal naming convention (UNC), for example, \\corporation\marketing.

Beginning with Visio 2002, you can save your drawing as an XML drawing (.vdx), an XML stencil (.vsx), or an XML template (.vtx).

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this method maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVDocument.SaveAs(string)**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SaveAs** method. Before running this macro, change path to the location where you want to save the drawing, and change filename to the name you'd like to assign the file.


```vb
 
Public Sub SaveAs_Example()  
  
    'Use the SaveAs method to save a document for the first time.  
    ThisDocument.SaveAs "path\filename .vsd" 
  
End Sub
```


