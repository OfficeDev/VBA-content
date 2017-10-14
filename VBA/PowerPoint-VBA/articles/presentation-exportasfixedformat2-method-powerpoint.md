---
title: Presentation.ExportAsFixedFormat2 Method (PowerPoint)
keywords: vbapp10.chm583126
f1_keywords:
- vbapp10.chm583126
ms.assetid: b1101e58-e6a8-9dd4-7071-1325ba71edb1
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Presentation.ExportAsFixedFormat2 Method (PowerPoint)

Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.


## Syntax

 _expression_. **ExportAsFixedFormat2**_(Path,_ _FixedFormatType,_ _Intent,_ _FrameSlides,_ _HandoutOrder,_ _OutputType,_ _PrintHiddenSlides,_ _PrintRange,_ _RangeType,_ _SlideShowName,_ _IncludeDocProperties,_ _KeepIRMSettings,_ _DocStructureTags,_ _BitmapMissingFonts,_ _UseISO19005_1,_ _IncludeMarkup,_ _ExternalExporter)_

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path for the export.|
| _FixedFormatType_|Required|**PpFixedFormatType**|The format to which the slides should be exported.|
| _Intent_|Optional|**PpFixedFormatIntent**|The purpose of the export.|
| _FrameSlides_|Optional|**MsoTriState**|Whether the slides to be exported should be bordered by a frame.|
| _HandoutOrder_|Optional|**PpPrintHandoutOrder**|The order in which the handout should be printed.|
| _OutputType_|Optional|**PpPrintOutputType**|The type of output.|
| _PrintHiddenSlides_|Optional|**MsoTriState**|Whether to print hidden slides.|
| _PrintRange_|Optional|**PrintRange**|The slide range.|
| _RangeType_|Optional|**PpPrintRangeType**|The type of slide range.|
| _SlideShowName_|Optional|**String**|The name of the slide show.|
| _IncludeDocProperties_|Optional|**Boolean**|Whether the document properties should also be exported. The default is  **False**.|
| _KeepIRMSettings_|Optional|**Boolean**|Whether the IRM settings should also be exported. The default is  **True**.|
| _DocStructureTags_|Optional|**Boolean**|Whether to include document structure tags to improve document accessibility. The default is  **True**.|
| _BitmapMissingFonts_|Optional|**Boolean**|Whether to include a bitmap of the text. The default is  **True**.|
| _UseISO19005_1_|Optional|**Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is  **False**.|
| _IncludeMarkup_|Optional|**Boolean**|Whether the resulting document should include associated pen marks.|
| _ExternalExporter_|Optional|**Variant**|A pointer to an Office add-in that implements the  **IMsoDocExporter** COM interface and allows calls to an alternate implementation of code. The default is a null pointer.|
| _Path_|Required|STRING||
| _FixedFormatType_|Required|PPFIXEDFORMATTYPE||
| _Intent_|Optional|PPFIXEDFORMATINTENT||
| _FrameSlides_|Optional|<unknown||
| _HandoutOrder_|Optional|PPPRINTHANDOUTORDER||
| _OutputType_|Optional|PPPRINTOUTPUTTYPE||
| _PrintHiddenSlides_|Optional|<unknown||
| _PrintRange_|Optional|PRINTRANGE||
| _RangeType_|Optional|PPPRINTRANGETYPE||
| _SlideShowName_|Optional|STRING||
| _IncludeDocProperties_|Optional|BOOL||
| _KeepIRMSettings_|Optional|BOOL||
| _DocStructureTags_|Optional|BOOL||
| _BitmapMissingFonts_|Optional|BOOL||
| _UseISO19005_1_|Optional|BOOL||
| _IncludeMarkup_|Optional|BOOL||
| _ExternalExporter_|Optional|VARIANT||

### Return value

 **VOID**


