---
title: Shapes.AddTextEffect Method (Publisher)
keywords: vbapb10.chm2162721
f1_keywords:
- vbapb10.chm2162721
ms.prod: publisher
api_name:
- Publisher.Shapes.AddTextEffect
ms.assetid: 21af82f1-d507-3c16-72df-bde1b5e00717
ms.date: 06/08/2017
---


# Shapes.AddTextEffect Method (Publisher)

Adds a new  **Shape** object representing a WordArt object to the specified **Shapes** collection.


## Syntax

 _expression_. **AddTextEffect**( **_PresetTextEffect_**,  **_Text_**,  **_FontName_**,  **_FontSize_**,  **_FontBold_**,  **_FontItalic_**,  **_Left_**,  **_Top_**)

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PresetTextEffect|Required| **MsoPresetTextEffect**|The preset text effect to use. The values of the  **MsoPresetTextEffect** constants correspond to the formats listed in the **WordArt Gallery** dialog box (numbered from left to right and from top to bottom).|
|Text|Required| **String**|The text to use for the WordArt object.|
|FontName|Required| **String**|The name of the font to use for the WordArt object.|
|FontSize|Required| **Variant**|The font size to use for the WordArt object. Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|
|FontBold|Required| **MsoTriState**|Determines whether to format the WordArt text as bold.|
|FontItalic|Required| **MsoTriState**|Determines whether to format the WordArt text as italic.|
|Left|Required| **Variant**|The position of the left edge of the shape representing the WordArt object.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the WordArt object.|

### Return Value

Shape


## Remarks

For the Left and Top parameters, numeric values are evaluated in points; strings can be in any units supported by Publisher (for example, "2.5 in").

The height and width of the WordArt object is determined by its text and formatting.

Use the  **[TextEffect](shape-texteffect-property-publisher.md)** property to return a **[TextEffectFormat](texteffectformat-object-publisher.md)** object whose properties can be used to edit an existing WordArt object.

The PresetTextEffect parameter can be one of the  ** [MsoPresetTextEffect](http://msdn.microsoft.com/library/56a7008d-ce2c-f127-56de-851cb8fef44f%28Office.15%29.aspx)** constants declared in the Microsoft Office type library. The **msoTextEffectMixed** constant is not supported.

The FontBold parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|Do not format the WordArt text as bold.|
| **msoTrue**|Format the WordArt text as bold.|
The FontItalic parameter can be one of the  **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Do not format the WordArt text as italic.|
| **msoTrue**|Format the WordArt text as italic.|

## Example

The following example adds a WordArt object to the first page of the active publication.


```vb
Dim shpWordArt As Shape 
 
Set shpWordArt = ActiveDocument.Pages(1).Shapes.AddTextEffect _ 
 (PresetTextEffect:=msoTextEffect7, Text:="Annual Report", _ 
 FontName:="Arial Black", FontSize:=24, _ 
 FontBold:=msoFalse, FontItalic:=msoFalse, _ 
 Left:=144, Top:=72) 

```


