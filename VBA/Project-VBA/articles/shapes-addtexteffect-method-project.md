---
title: Shapes.AddTextEffect Method (Project)
ms.prod: project-server
ms.assetid: 5510367c-7f8d-3266-642f-61f3d45a18cf
ms.date: 06/08/2017
---


# Shapes.AddTextEffect Method (Project)
The  **AddTextEffect** method is not implemented in Project.

## Syntax

 _expression_. **AddTextEffect** _(PresetTextEffect,_ _Text,_ _FontName,_ _FontSize,_ _FontBold,_ _FontItalic,_ _Left,_ _Top)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PresetTextEffect_|Required|**MsoPresetTextEffect**|A preset text effect. The values of the  **MsoPresetTextEffect** constants correspond to the formats listed in the WordArt Gallery dialog box (numbered from left to right and from top to bottom).|
| _Text_|Required|**String**|The text in the WordArt.|
| _FontName_|Required|**String**|The name of the font used in the WordArt.|
| _FontSize_|Required|**Single**|The size (in points) of the font used in the WordArt.|
| _FontBold_|Required|**MsoTriState**|Use the  **msoTrue** constant to bold the font; otherwise, use **msoFalse**.|
| _FontItalic_|Required|**MsoTriState**|Use the  **msoTrue** constant to italicize the font; otherwise, use **msoFalse**.|
| _Left_|Required|**Single**|The position, in points, of the left edge of the WordArt shape relative to the left edge of the report.|
| _Top_|Required|**Single**|The position, in points, of the top edge of the WordArt shape relative to the top edge of the report.|
| _PresetTextEffect_|Required|MSOPRESETTEXTEFFECT||
| _Text_|Required|STRING||
| _FontName_|Required|STRING||
| _FontSize_|Required|FLOAT||
| _FontBold_|Required|MSOTRISTATE||
| _FontItalic_|Required|MSOTRISTATE||
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
|Name|Required/Optional|Data type|Description|

### Return value

 **Shape**


## Remarks


 **Note**  The  **Shapes.AddTextEffect** method in Excel and Word creates a WordArt item, and returns a **Shape** object that represents the new WordArt item. But, Project does not support directly creating a WordArt item.

Instead of using the  **AddTextEffect** method to add WordArt, you can use **AddTextbox**, and then format the selected text box with WordArt styles.


## See also


#### Other resources


[Shapes Object](shapes-object-project.md)
[Shape Object](shape-object-project.md)
[MsoPresetTextEffect Enumeration](http://msdn.microsoft.com/en-us/library/office/ff861792%28v=office.15%29)
