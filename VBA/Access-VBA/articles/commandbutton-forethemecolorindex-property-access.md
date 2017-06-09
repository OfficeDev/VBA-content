---
title: CommandButton.ForeThemeColorIndex Property (Access)
keywords: vbaac10.chm14604
f1_keywords:
- vbaac10.chm14604
ms.prod: access
api_name:
- Access.CommandButton.ForeThemeColorIndex
ms.assetid: 4831634a-6988-57ec-0e47-6c16a6c832a0
ms.date: 06/08/2017
---


# CommandButton.ForeThemeColorIndex Property (Access)

Gets or sets a value that represents a color in the applied color theme associated with the  **ForeColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **ForeThemeColorIndex**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **ForeThemeColorIndex** property contains one of the index values listed in the following table.



|**Index Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **ForeThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## Example

The following code example sets the Fore Color to the Text 2 color by setting the  **ForeThemeColorIndex** property.


```vb
Me.ctl.ForeThemeColorIndex=2
```


## See also


#### Concepts


[CommandButton Object](commandbutton-object-access.md)

