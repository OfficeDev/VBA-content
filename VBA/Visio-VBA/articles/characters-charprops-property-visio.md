---
title: Characters.CharProps Property (Visio)
keywords: vis_sdr.chm10213225
f1_keywords:
- vis_sdr.chm10213225
ms.prod: visio
api_name:
- Visio.Characters.CharProps
ms.assetid: 7c05633d-9e99-cee3-0d24-bff6d191ef24
ms.date: 06/08/2017
---


# Characters.CharProps Property (Visio)

Sets a character property of a  **Characters** object to a new value. Write-only.


## Syntax

 _expression_ . **CharProps**( **_CellIndex_** )

 _expression_ An expression that returns a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CellIndex_|Required| **Integer**|The index of the cell in the Character section to set. See Remarks for possible values.|

### Return Value

Integer


## Remarks

Depending on the extent of the text range and the format, setting the  **CharProps** property may cause rows to be added or removed from a shape's Character ShapeSheet section.

The  **CharProps** property is a write-only property. To retrieve formatting properties of a **Characters** object, use the **CharPropsRow** property.

The values of the CellIndex argument correspond to cells in the Character section of the ShapeSheet window, and the values of the  **CharProps** property correspond to the values that can be entered in those cells.

Constants for CellIndex and for the  **CharProps** property value are declared in the Visio type library in **VisCellIndices** .



|** CellIndex**|** Value**|** intExpression**|** Value**|
|:-----|:-----|:-----|:-----|
| **visCharacterFont**| 0| An integer that represents an index into the **Fonts** collection installed on a system. Zero (0) represents the default font.| N/A|
| **visCharacterColor**| 1| An integer from 0 to 23 that corresponds to a color in the current color palette.| N/A|
| **visCharacterStyle**| 2| **visBold** **visItalic** **visUnderLine** **visSmallCaps**| &;H1 &;H2 &;H4 &;H8|
| **visCharacterCase**| 3| **visCaseNormal** **visCaseAllCaps** **visCaseInitialCaps**| 0 1 2|
| **visCharacterPos**| 4| **visPosNormal** **visPosSuper** **visPosSub**| 0 1 2|
| **visCharacterSize**| 7| An integer that represents point size.| N/A|
| **visCharacterColorTrans**|17|An integer from 0 to 100 that corresponds to the degree of transparency of the text color, as a percentage.|N/A|
| **visCharacterDblUnderline**|8| **Boolean**|N/A|
| **visCharacterFontScale**|5|An integer from 0 to 655 that represents the width of the text font, as a percentage, relative to the default (100%). |N/A|
| **visCharacterLangID**|57|A  **Long** that represents the language the text is in. The language ID (LANGID) for a character is a 16-bit value defined by Windows, consisting of a primary language ID and a secondary language ID. To determine the value for particular languages, see the Platform SDK documentation on MSDN.|N/A|
| **visCharacterLetterspace**|16|An integer that represents additional space between adjacent letters, in points.|N/A|
| **visCharacterOverline**|9| **Boolean**|N/A|
| **visCharacterStrikethru**|10| **Boolean**|N/A|
If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters.set_CharProps**
    

