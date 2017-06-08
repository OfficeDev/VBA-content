---
title: Presentation.SaveCopyAs Method (PowerPoint)
keywords: vbapp10.chm583037
f1_keywords:
- vbapp10.chm583037
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.SaveCopyAs
ms.assetid: 456415d1-845a-9e9b-45ce-98985e94aee5
ms.date: 06/08/2017
---


# Presentation.SaveCopyAs Method (PowerPoint)

Saves a copy of the specified presentation to a file without modifying the original.


## Syntax

 _expression_. **SaveCopyAs**( **_FileName_**, **_FileFormat_**, **_EmbedTrueTypeFonts_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the name to save the file under. If you don't include a full path, PowerPoint saves the file in the current folder.|
| _FileFormat_|Optional|**PpSaveAsFileType**|The file format.|
| _EmbedTrueTypeFonts_|Optional|**MsoTriState**|Specifies whether TrueType fonts are embedded.|

## Remarks

The  _FileFormat_ parameter value can be one of these **PpSaveAsFileType** constants. The default is **ppSaveAsDefault**.


||
|:-----|
|**ppSaveAsHTMLv3**|
|**ppSaveAsAddIn**|
|**ppSaveAsBMP**|
|**ppSaveAsDefault**|
|**ppSaveAsGIF**|
|**ppSaveAsHTML**|
|**ppSaveAsHTMLDual**|
|**ppSaveAsJPG**|
|**ppSaveAsMetaFile**|
|**ppSaveAsPNG**|
|**ppSaveAsPowerPoint3**|
|**ppSaveAsPowerPoint4**|
|**ppSaveAsPowerPoint4FarEast**|
|**ppSaveAsPowerPoint7**|
|**ppSaveAsPresentation**|
|**ppSaveAsRTF**|
|**ppSaveAsShow**|
|**ppSaveAsTemplate**|
|**ppSaveAsTIF**|
|**ppSaveAsWebArchive**|
The  _EmbedTrueTypeFonts_ parameter value can be one of these **MsoTriState** constants.



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

