---
title: WebOptions.RelyOnCSS Property (Word)
keywords: vbawd10.chm165937155
f1_keywords:
- vbawd10.chm165937155
ms.prod: word
api_name:
- Word.WebOptions.RelyOnCSS
ms.assetid: 7d203567-e4a8-ac5a-4307-739d07996c2d
ms.date: 06/08/2017
---


# WebOptions.RelyOnCSS Property (Word)

 **True** if cascading style sheets (CSS) are used for font formatting when you view a saved document in a Web browser. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **RelyOnCSS**

 _expression_ Required. A variable that represents a **[WebOptions](weboptions-object-word.md)** collection.


## Remarks

Microsoft Word creates a cascading style sheet file and saves it either to the specified folder or to the same folder as your Web page, depending on the value of the  **[OrganizeInFolder](weboptions-organizeinfolder-property-word.md)** property. **False** if HTML <FONT> tags and cascading style sheets are used.

You should set this property to  **True** if your Web browser supports cascading style sheets because this will give you more precise layout and formatting control on your Web page and make it look more like your document (as it appears in Microsoft Word).


## Example

This example enables the use of cascading style sheets as the global default for the application.


```vb
Application.DefaultWebOptions.RelyOnCSS = True
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-word.md)

