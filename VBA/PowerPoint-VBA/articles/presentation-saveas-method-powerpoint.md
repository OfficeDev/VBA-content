---
title: Presentation.SaveAs Method (PowerPoint)
keywords: vbapp10.chm583036
f1_keywords:
- vbapp10.chm583036
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SaveAs
ms.assetid: d70a678b-66ed-9dd6-5a5e-454cdf808784
ms.date: 06/08/2017
---

# Presentation.SaveAs Method (PowerPoint)

Saves a presentation that's never been saved, or saves a previously saved presentation under a different name.


## Syntax

 _expression_. **SaveAs**( **_Filename_**, **_FileFormat_**, **_EmbedFonts_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required|**String**|Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.|
| _FileFormat_|Optional|**PpSaveAsFileType**|Specifies the saved file format. If this argument is omitted, the file is saved in the default file format ( **ppSaveAsDefault** ).|
| _EmbedFonts_|Optional|**MsoTriState**|Specifies whether PowerPoint embeds TrueType fonts in the saved presentation.|

## Remarks

The  _FileFormat_ parameter value can be one of these **PpSaveAsFileType** constants. The default is **ppSaveAsDefault**. For a complete list of constants, see [PpSaveAsFileType Enumeration](https://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint.ppsaveasfiletype.aspx).

||
|:-----|
|**ppSaveAsAddIn**|
|**ppSaveAsBMP**|
|**ppSaveAsDefault**|
|**ppSaveAsEMF**|
|**ppSaveAsExternalConverter**|
|**ppSaveAsGIF**|
|**ppSaveAsJPG**|
|**ppSaveAsMetaFile**|
|**ppSaveAsMP4**|
|**ppSaveAsOpenDocumentPresentation**|
|**ppSaveAsOpenXMLAddin**|
|**ppSaveAsOpenXMLPicturePresentation**|
|**ppSaveAsOpenXMLPresentation**|
|**ppSaveAsOpenXMLPresentationMacroEnabled**|
|**ppSaveAsOpenXMLShow**|
|**ppSaveAsOpenXMLShowMacroEnabled**|
|**ppSaveAsOpenXMLTemplate**|
|**ppSaveAsOpenXMLTemplateMacroEnabled**|
|**ppSaveAsOpenXMLTheme**|
|**ppSaveAsPDF**|
|**ppSaveAsPNG**|
|**ppSaveAsPresentation**|
|**ppSaveAsRTF**|
|**ppSaveAsShow**|
|**ppSaveAsStrictOpenXMLPresentation**|
|**ppSaveAsTemplate**|
|**ppSaveAsTIF**|
|**ppSaveAsWMV**|
|**ppSaveAsXMLPresentation**|
|**ppSaveAsXPS**|
The  _EmbedFonts_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|TrueType fonts are not embedded.|
|**msoTriStateMixed**|Embedded fonts are a mixture of TrueType and non-TrueType. The default. |
|**msoTrue**|TrueType fonts are embedded.|

## Example

This example saves a copy of the active presentation under the name "New Format Copy.ppt." By default, this copy is saved in the format of a presentation in the current version of PowerPoint. The presentation is then saved as a PowerPoint 4.0 file named "Old Format Copy."


```vb
With Application.ActivePresentation 
    .SaveCopyAs "New Format Copy" 
    .SaveAs "Old Format Copy", ppSaveAsPowerPoint4 
End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

