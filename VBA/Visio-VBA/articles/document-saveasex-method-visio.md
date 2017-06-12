---
title: Document.SaveAsEx Method (Visio)
keywords: vis_sdr.chm10516505
f1_keywords:
- vis_sdr.chm10516505
ms.prod: visio
api_name:
- Visio.Document.SaveAsEx
ms.assetid: c0bef38d-1849-67ab-606f-8178de46c7c6
ms.date: 06/08/2017
---


# Document.SaveAsEx Method (Visio)

Saves a document with a file name using extra information passed in an argument.


## Syntax

 _expression_ . **SaveAsEx**( **_FileName_** , **_SaveFlags_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The file name for the document.|
| _SaveFlags_|Required| **Integer**|How to save the file.|

### Return Value

Nothing


## Remarks

The  **SaveAsEx** method is identical to the **SaveAs** method, except that it provides an extra argument in which the caller can specify how the document is to be saved.

The  _SaveFlags_ argument should be a combination of the following values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSaveAsRO**|&;H1|The document is saved as read-only.|
| **visSaveAsWS**|&;H2|The current workspace is saved with the file.|
| **visSaveAsListInMRU**|&;H4|The document is included in the Most Recently Used (MRU) list. By default,  **Save** and **SaveAs** do not place the document into the MRU list.|

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **SaveAsEx** method. Before running this macro, change _path_ to the location where you want to save the drawing, and change _filename_ to the name you'd like to assign the file.


```vb
 
Public Sub SaveAsEx_Example() 
 
 'Use the SaveAsEx method to save the drawing as a 
 'new read-only drawing. 
 ThisDocument.SaveAsEx "path\filename.vsd ", visSaveAsRO 
 
End Sub
```


