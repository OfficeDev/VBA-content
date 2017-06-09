---
title: Application.SmartArtQuickStyles Property (Word)
keywords: vbawd10.chm158335458
f1_keywords:
- vbawd10.chm158335458
ms.prod: word
api_name:
- Word.Application.SmartArtQuickStyles
ms.assetid: 47cca923-fc88-6973-926c-2fa69c2f0f10
ms.date: 06/08/2017
---


# Application.SmartArtQuickStyles Property (Word)

Returns a [SmartArtQuickStyles](http://msdn.microsoft.com/library/d488ac12-160b-c518-2b56-cc0a3a45c6b7%28Office.15%29.aspx) object that represents the set of SmartArt styles that are currently loaded in the application. Read-only.


## Syntax

 _expression_ . **SmartArtQuickStyles**

 _expression_ An expression that returns a **[Application](application-object-word.md)** object.


## Remarks

The set of styles represented by the  **SmartArtQuickStyles** property correspond to the available styles in the **Styles** group on the **Design tab** on the **SmartArt Tools** contextual tab in Word.


## Example

The following code example adds a SmartArt graphic to the active document and then sets the SmartArt graphic style to "Polished".


```vb
Dim myShape As Shape 
Dim mySmartArt As SmartArt 
 
Set myShape = ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts(1), 50, 50, 200, 200) 
Set mySmartArt = myShape.SmartArt 
 
mySmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(6)
```


## See also


#### Concepts


[Application Object](application-object-word.md)

