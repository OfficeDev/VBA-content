---
title: Document.SaveAs Method (Publisher)
keywords: vbapb10.chm196696
f1_keywords:
- vbapb10.chm196696
ms.prod: publisher
api_name:
- Publisher.Document.SaveAs
ms.assetid: ba8b85d7-8ca9-dcf5-12b4-4cabced743e6
ms.date: 06/08/2017
---


# Document.SaveAs Method (Publisher)

Saves the specified publication with a new name or format.


## Syntax

 _expression_. **SaveAs**( **_Filename_**,  **_Format_**,  **_AddToRecentFiles_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Filename|Optional| **Variant**|The name for the publication. The default is the current folder and file name. If the publication has never been saved, the default name is used, for example, Publication1.pub. If a publication with the specified file name already exists, the publication is overwritten without the user being prompted first.|
|Format|Optional| **PbFileFormat**|The format in which the publication is saved.|
|AddToRecentFiles|Optional| **Boolean**| **True** to add the publication to the list of recently used files on the File menu. Default is **True**.|

## Remarks

The Format parameter can be one of the  **PbFileFormat** constants declared in the Microsoft Publisher type library and shown in the following table. The default is **pbFilePublication**.



| **pbFileHTMLFiltered**|
| **pbFilePublication**|
| **pbFilePublicationHTML**|
| **pbFilePublisher2000**|
| **pbFilePublisher98**|
| **pbFileRTF**|
| **pbFileWebArchive**|
If there is insufficient memory or disk space to save the file, an error occurs.

Calling the  **SaveAs** method always performs saves in the foreground regardless of whether background saves are enabled.


## Example

This example saves the active publication as a Microsoft Publisher 2000 file.


```vb
ActiveDocument.SaveAs FileName:="ReportPub2000.pub", Format:=pbFilePublisher2000
```

This example saves the active publication as Test.rtf in Rich Text Format (RTF).




```vb
ActiveDocument.SaveAs FileName:="Test.rtf", Format:=pbFileRTF
```

This example saves the active Web publication as a set of filtered HTML pages and supporting files. Note that the .htm extension is automatically added to the value of the Filename parameter if the value of the Format parameter is  **pbFileHTMLFiltered**.




```vb
With ActiveDocument 
 .SaveAs Filename:="CompanyContacts", Format:=pbFileHTMLFiltered 
End With
```


