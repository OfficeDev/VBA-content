---
title: Characters.ParaProps Property (Visio)
keywords: vis_sdr.chm10214030
f1_keywords:
- vis_sdr.chm10214030
ms.prod: visio
api_name:
- Visio.Characters.ParaProps
ms.assetid: 8f71a7ba-3a9e-01b4-1bbe-040fd441a284
ms.date: 06/08/2017
---


# Characters.ParaProps Property (Visio)

Sets the paragraph property of a  **Characters** object to a new value. Read/write.


## Syntax

 _expression_ . **ParaProps**( **_CellIndex_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CellIndex_|Required| **Integer**|The cell value to set.|

### Return Value

Integer


## Remarks

The possible values of the CellIndex argument correspond to named cells in the Paragraph section of the ShapeSheet. Constants for CellIndex are declared by the Visio type library in  **VisCellIndices** .



|** Constant**|** Value**|
|:-----|:-----|
| **visIndentFirst**| 0|
| **visIndentLeft**| 1|
| **visIndentRight**| 2|
| **visSpaceLine**| 3|
| **visSpaceBefore**| 4|
| **visSpaceAfter**| 5|
| **visHorzAlign**| 6|
| **visBulletIndex**| 7|
Depending on the extent of the text range and the format, setting the  **ParaProps** property may cause rows to be added or removed from the Paragraph section of the ShapeSheet.

To retrieve information about an existing format, use the  **ParaPropsRow** property.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this property maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters.set_ParaProps**
    

