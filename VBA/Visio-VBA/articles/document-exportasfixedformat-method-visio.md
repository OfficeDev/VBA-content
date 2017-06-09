---
title: Document.ExportAsFixedFormat Method (Visio)
keywords: vis_sdr.chm10560095
f1_keywords:
- vis_sdr.chm10560095
ms.prod: visio
api_name:
- Visio.Document.ExportAsFixedFormat
ms.assetid: 70b83f7e-b7f8-7b8f-d9d7-7f7b30f3b45d
ms.date: 06/08/2017
---


# Document.ExportAsFixedFormat Method (Visio)

Exports a Microsoft Visio document as a file in a fixed format, either PDF or XPS.


## Syntax

 _expression_ . **ExportAsFixedFormat**( **_FixedFormat_** , **_OutputFileName_** , **_Intent_** , **_PrintRange_** , **_FromPage_** , **_ToPage_** , **_ColorAsBlack_** , **_IncludeBackground_** , **_IncludeDocumentProperties_** , **_IncludeStructureTags_** , **_UseISO19005_1_** , **_FixedFormatExtClass_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FixedFormat_|Required| **VisFixedFormatTypes**|The format type in which to export the document. See Remarks for possible values.|
| _OutputFileName_|Optional| **String**|The name and path of the file to which to output, enclosed in quotation marks.|
| _Intent_|Required| **VisDocExIntent**|The output quality. See Remarks for possible values.|
| _PrintRange_|Required| **VisPrintOutRange**|The range of document pages to be exported. See Remarks for possible values.|
| _FromPage_|Optional| **Long**| If _PrintRange_ is **visPrintFromTo** , the first page in the range to be exported. The default is 1, which indicates the first page of the drawing.|
| _ToPage_|Optional| **Long**|If  _PrintRange_ is **visPrintFromTo** , the last page in the range to be exported. The default is -1, which indicates the last page of the drawing.|
| _ColorAsBlack_|Optional| **Boolean**| **True** to render all colors as black to ensure that all shapes are visible in the exported drawing. **False** to render colors normally. The default is **False** .|
| _IncludeBackground_|Optional| **Boolean**|Whether to include background pages in the exported file. The default is  **True** .|
| _IncludeDocumentProperties_|Optional| **Boolean**|Whether to include document properties in the exported file. The default is  **True** .|
| _IncludeStructureTags_|Optional| **Boolean**|Whether to include document structure tags to improve document accessibility. The default is  **True** .|
| _UseISO19005_1_|Optional| **Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is  **False** .|
| _FixedFormatExtClass_|Optional| **[UNKNOWN]**|A pointer to a class that implements the  **IMsoDocExporter** interface for purposes of creating custom fixed output. The default is a null pointer.|

### Return Value

Nothing


## Remarks

The  **ExportAsFixedFormat** method creates a file that contains a static view of the Visio document.

Possible values for the  _FixedFormat_ parameter are shown in the following table and declared in **VisFixedFormatTypes** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visFixedFormatPDF**|1|PDF fixed format|
| **visFixedFormatXPS**|2|XPS fixed format|
Possible values for the  _Intent_ parameter are shown in the following table and declared in **VisDocExIntent** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visDocExIntentPrint**|1|Intended to be published online and printed|
| **visDocExIntentScreen**|0|Intended to be published only online|
Possible values for the  _PrintRange_ parameter are shown in the following table and declared in **VisPrintOutRange** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPrintAll**|0|Prints all foreground pages.|
| **visPrintCurrentPage**|2|Prints the active page.|
| **visPrintCurrentView**|4|Prints the current view area.|
| **visPrintFromTo**|1|Prints pages between the  _FromPage_ value and the _ToPage_ value.|
| **visPrintSelection**|3|Prints a selection.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **ExportAsFixedFormat** method to export the active Visio document to the root of the C drive in PDF format.


```vb
Public Sub ExportAsFixedFormat_Example() 
 
    ActiveDocument.ExportAsFixedFormat visFixedFormatPDF, "C:\ExportedVisioDocument .pdf", visDocExIntentPrint, visPrintAll 
 
End Sub
```


