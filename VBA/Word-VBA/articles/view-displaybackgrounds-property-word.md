---
title: View.DisplayBackgrounds Property (Word)
keywords: vbawd10.chm161808433
f1_keywords:
- vbawd10.chm161808433
ms.prod: word
api_name:
- Word.View.DisplayBackgrounds
ms.assetid: 6b1dfa3a-a2bd-a737-e0b2-0792d13451ba
ms.date: 06/08/2017
---


# View.DisplayBackgrounds Property (Word)

Returns or sets a  **Boolean** that represents whether background colors and images are shown when a document is displayed in print layout view. .


## Syntax

 _expression_ . **DisplayBackgrounds**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Remarks

Corresponds to the  **Background colors and images (Print view only)** option located on the **View** tab of the **Options** dialog box.


## Example

The following example hides background colors and images when the active document is displayed in print layout view.


```vb
ActiveDocument.ActiveWindow.View.DisplayBackgrounds = False
```


## See also


#### Concepts


[View Object](view-object-word.md)

