---
title: EmailOptions.RelyOnCSS Property (Word)
keywords: vbawd10.chm165347634
f1_keywords:
- vbawd10.chm165347634
ms.prod: word
api_name:
- Word.EmailOptions.RelyOnCSS
ms.assetid: 4c1d89df-cb96-2463-e884-3f6571dfff6e
ms.date: 06/08/2017
---


# EmailOptions.RelyOnCSS Property (Word)

 **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnCSS**

 _expression_ Required. A variable that represents an **[EmailOptions](emailoptions-object-word.md)** collection.


## Remarks

Microsoft Word creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the  **OrganizeInFolder** property of the **[WebOptions](weboptions-object-word.md)** or **[DefaultWebOptions](defaultweboptions-object-word.md)** object. This property returns **False** if HTML <FONT> tags and cascading style sheets are used. The default value is **True** .

You should set this property to  **True** if your Web browser supports cascading style sheets because this will give you more precise layout and formatting control on your Web page and make it look more like your document (as it appears in Microsoft Word).


## Example

This example enables the use of cascading style sheets as the global default for the application.


```vb
Application.DefaultWebOptions.RelyOnCSS = True
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

