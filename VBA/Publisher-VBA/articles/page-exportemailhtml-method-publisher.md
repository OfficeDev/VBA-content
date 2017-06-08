---
title: Page.ExportEmailHTML Method (Publisher)
keywords: vbapb10.chm393273
f1_keywords:
- vbapb10.chm393273
ms.prod: publisher
api_name:
- Publisher.Page.ExportEmailHTML
ms.assetid: 6257e9b5-26b5-73ae-7d40-50dd0a764488
ms.date: 06/08/2017
---


# Page.ExportEmailHTML Method (Publisher)

Exports the active page of the publication as an HTML file.


## Syntax

 _expression_. **ExportEmailHTML**( **_Filename_**)

 _expression_A variable that represents a  **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Required| **String**|The name of the file to which to export the HTML.|

## Remarks

If the name of an existing HTML file is specified, that file is overwritten.

This method can only be used on the active page of the publication.


## Example

The following example sets the first page in the document as the active page, and exports that page to a file. (Note that PathToFile must be replaced with a valid file path for this example to work.)


```vb
Sub ExportEmail() 
 Dim strFilePath As String 
 strFilePath = "PathToFile" 
 With ActiveDocument.ActiveView 
 .ActivePage = ActiveDocument.Pages(1) 
 .ActivePage.ExportEmailHTML (strFilePath) 
 End With 
 
End Sub
```


