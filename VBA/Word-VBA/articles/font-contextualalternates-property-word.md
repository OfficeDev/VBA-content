---
title: Font.ContextualAlternates Property (Word)
keywords: vbawd10.chm156369073
f1_keywords:
- vbawd10.chm156369073
ms.prod: word
api_name:
- Word.Font.ContextualAlternates
ms.assetid: 065589b0-afbd-dfb1-4f96-2c70b558b773
ms.date: 06/08/2017
---


# Font.ContextualAlternates Property (Word)

Specifies whether or not contextual alternates are enabled for the specified font. Read/write  **Long** .


## Syntax

 _expression_ . **ContextualAlternates**

 _expression_ An expression that returns a **[Font](font-object-word.md)** object.


## Remarks

Contextual alternates are ligatures that are applied to individual characters based on the letters around them (their context). Contextual alternates can also be applied to entire words in certain contexts, for example, words frequently used in titles (such as "of" and "the"). When contextual alternates are enabled for a font, they are used instead of the standard ligatures in those contexts defined by the font designer.

Setting this property has the same effect as selecting the check box next to  **Use Contextual Alternates** (in the **OpenType Features** group, **Advanced** tab, on the **Font** dialog in Word).


## Example

The following code example enables contextual alternates for the font in the active document.


```vb
ActiveDocument.Range.Font.ContextualAlternates = True
```


## See also


#### Concepts


[Font Object](font-object-word.md)

