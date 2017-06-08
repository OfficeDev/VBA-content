---
title: Shapes.AddWordArt Method (Publisher)
keywords: vbapb10.chm2162761
f1_keywords:
- vbapb10.chm2162761
ms.prod: publisher
api_name:
- Publisher.Shapes.AddWordArt
ms.assetid: 8ff83baa-5d88-5f80-3a69-5f712ba5e583
ms.date: 06/08/2017
---


# Shapes.AddWordArt Method (Publisher)

Returns a  **Shape** object that represents the WordArt to be added to the publication.


## Syntax

 _expression_. **AddWordArt**( **_PresetWordArt_**,  **_Text_**,  **_FontName_**,  **_FontSize_**,  **_FontBold_**,  **_FontItalic_**,  **_Left_**,  **_Top_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PresetWordArt|Required| **pbPresetWordArt**|The type of preset WordArt to add.|
|Text|Required| **String**|The text of the WordArt.|
|FontName|Required| **String**|The name of the font to be used in the WordArt.|
|FontSize|Required| **Variant**|The size of the font to be used in the WordArt.|
|FontBold|Required| **[MSOTRISTATE]**|Whether the WordArt text should be bold. See Remarks for possible values.|
|FontItalic|Required| **[MSOTRISTATE]**|Whether the WordArt text should be italic. See Remarks for possible values.|
|Left|Required| **Variant**|The horizontal position of the WordArt.|
|Top|Required| **Variant**|The vertical position of the WordArt.|

### Return Value

 **Shape**


### Remarks

The  **FontBold** parameter value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the WordArt are formatted as bold.|
| **msoTriStateMixed**|Return value indicating that the WordArt contains some text formatted as bold and some text not formatted as bold.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the WordArt are formatted as bold.|
The  **FontItalic** parameter value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the WordArt are formatted as italic.|
| **msoTriStateMixed**|Return value indicating that the WordArt contains some text formatted as italic and some text not formatted as italic.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the WordArt are formatted as italic.|

