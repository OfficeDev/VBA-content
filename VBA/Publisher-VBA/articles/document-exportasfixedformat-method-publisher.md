---
title: Document.ExportAsFixedFormat Method (Publisher)
keywords: vbapb10.chm196758
f1_keywords:
- vbapb10.chm196758
ms.prod: publisher
api_name:
- Publisher.Document.ExportAsFixedFormat
ms.assetid: 8bb5b64f-57b2-cf87-344c-be1e2741a59c
ms.date: 06/08/2017
---


# Document.ExportAsFixedFormat Method (Publisher)

Saves a Microsoft Publisher publication in PDF or XPS format. The conversion readies the document to be sent to commercial presses, to copy shops, for desktop printing, or for electronic distribution.


## Syntax

 _expression_. **ExportAsFixedFormat**( **_Format_**,  **_Filename_**,  **_Intent_**,  **_IncludeDocumentProperties_**,  **_ColorDownsampleTarget_**,  **_ColorDownsampleThreshold_**,  **_OneBitDownsampleTarget_**,  **_OneBitDownsampleThreshold_**, **_From_**, **_To_**, **_Copies_**, **_Collate_**, **_PrintStyle_**, **_DocStructureTags_**, **_BitmapMissingFonts_**, **_UseISO19005_1_**, **_ExternalExporter_**)

 _expression_A variable that represents a  **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Format|Required| **PbFixedFormatType**|The format in which you want to export the publication. See Remarks for possible values.|
|Filename|Required| **String**|The file name for the exported file.|
|Intent|Optional| **PbFixedFormatIntent**|The output quality of the exported file. See Remarks for possible values.|
|IncludeDocumentProperties|Optional| **Boolean**| **True** if you want to save the document properties with the PDF file.|
|ColorDownsampleTarget|Optional| **Long**|The target for down-sampling of colored images. Measured in dots per inch. Must be greater than 96 |
|ColorDownsampleThreshold|Optional| **Long**|The threshold at or above which an image is down-sampled to the ColorDownsample target level.|
|OneBitDownsampleTarget|Optional| **Long**|The target for down-sampling of one-bit images.|
|OneBitDownsampleThreshold|Optional| **Long**|The threshold at or above which an image is down-sampled to the OneBitDownsample target level.|
|From|Optional| **Long**|The page number of the first page to export.|
|To|Optional| **Long**|The page number of the last page to export.|
|Copies|Optional| **Long**|The number of copies.|
|Collate|Optional| **Boolean**|Whether to collate the copies.|
|PrintStyle|Optional| **PbPrintStyle**|The style in which to print the exported file. See Remarks for possible values.|
|DocStructureTags|Optional| **Boolean**|Whether to include document structure tags to improve document accessibility. The default is  **True**.|
|BitmapMissingFonts|Optional| **Boolean**|Whether to include a bitmap of the text. Pass  **True** for this parameter when font licenses do not permit a font to be embedded in the PDF file. If you pass **False**, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is  **True**. |
|UseISO19005_1|Optional| **Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is  **False**.|
|ExternalExporter|Optional| **Variant**|A pointer to an add-in that allows calls to an alternate implementation of code. You can add support for additional fixed formats by writing a Microsoft Office add-in that implements the  **IMsoDocExporter** COM interface. For more information, see "Extending the Office (2007) Fixed-Format Export Feature" on MSDN.|

## Remarks

The  **ExportAsFixedFormat** method is the equivalent of the **Publish As PDF or XPS** command on the **File** menu in the Publisher user interface.

Possible values for the Format parameter are shown in the following table and declared in  **PbFixedFormatType** in the Publisher type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbFixedFormatTypePDF**|2|PDF format|
| **pbFixedFormatTypeXPS**|1|XPS format|
Possible values for the Intent parameter are shown in the following table and declared in  **PbFixedFormatIntent** in the Publisher type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **pbIntentMinimum**|1|Squeeze the publication to the smallest file size. This satisfies the on-screen viewing scenario where the publication is viewed on a computer monitor.|
| **pbIntentStandard**|2 |Distribute the publication as an e-mail message or from a Web site. Note that the user does not know how the publication will be viewed: on-screen or printed from a desktop printer. Both the desktop printing scenario and the on-screen viewing scenario must be met by this intent.|
| **pbIntentPrinting**|3|Print the publication on a desktop printer or at a copy store.|
| **pbIntentCommercial **|4|Submit the publication to a commercial press.|
Possible values for the PrintStyle parameter are declared in the  **[PbPrintStyle](pbprintstyle-enumeration-publisher.md)** enumeration in the Publisher type library. The default value depends on the value of the Intent parameter.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ExportAsFixedFormat** method to save the the active publication as a .pdf file.

Before running this code, replace  _pathandfilename.pdf_ with a valid file name and the path to a folder on your computer where you have permission to save files.




```vb
Public Sub ExportAsFixedFormat_Example() 
 
 ThisDocument.ExportAsFixedFormat pbFixedFormatTypePDF, "pathandfilename.pdf" 
 
End Sub
```


