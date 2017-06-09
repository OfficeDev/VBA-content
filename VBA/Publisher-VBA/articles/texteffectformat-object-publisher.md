---
title: TextEffectFormat Object (Publisher)
keywords: vbapb10.chm3801087
f1_keywords:
- vbapb10.chm3801087
ms.prod: publisher
api_name:
- Publisher.TextEffectFormat
ms.assetid: 672d0ef0-cbcd-05ef-9aa5-b986c7b045ac
ms.date: 06/08/2017
---


# TextEffectFormat Object (Publisher)

Contains properties and methods that apply to WordArt objects.
 


## Example

Use the  **TextEffect** property to return a **TextEffectFormat** object. The following example sets the font name and formatting for shape one on the first page of the active publication. For this example to work, shape one must be a WordArt object.
 

 

```
Sub FormatWordArt() 
 With ActiveDocument.Pages(1).Shapes(1).TextEffect 
 .FontName = "Courier New" 
 .FontBold = MsoTrue 
 .FontItalic = MsoTrue 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ToggleVerticalText](texteffectformat-toggleverticaltext-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Alignment](texteffectformat-alignment-property-publisher.md)|
|[Application](texteffectformat-application-property-publisher.md)|
|[FontBold](texteffectformat-fontbold-property-publisher.md)|
|[FontItalic](texteffectformat-fontitalic-property-publisher.md)|
|[FontName](texteffectformat-fontname-property-publisher.md)|
|[FontSize](texteffectformat-fontsize-property-publisher.md)|
|[KernedPairs](texteffectformat-kernedpairs-property-publisher.md)|
|[NormalizedHeight](texteffectformat-normalizedheight-property-publisher.md)|
|[Parent](texteffectformat-parent-property-publisher.md)|
|[PresetShape](texteffectformat-presetshape-property-publisher.md)|
|[PresetTextEffect](texteffectformat-presettexteffect-property-publisher.md)|
|[PresetWordArt](texteffectformat-presetwordart-property-publisher.md)|
|[RotatedChars](texteffectformat-rotatedchars-property-publisher.md)|
|[Text](texteffectformat-text-property-publisher.md)|
|[Tracking](texteffectformat-tracking-property-publisher.md)|

