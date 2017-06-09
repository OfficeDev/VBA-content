---
title: Page.Caption Property (Access)
keywords: vbaac10.chm12148
f1_keywords:
- vbaac10.chm12148
ms.prod: access
api_name:
- Access.Page.Caption
ms.assetid: 7f1b5038-4543-c373-96e9-135102cdd6e6
ms.date: 06/08/2017
---


# Page.Caption Property (Access)

Gets or sets the text that appears at the top of the page. Read/write  **String**.


## Syntax

 _expression_. **Caption**

 _expression_ A variable that represents a **Page** object.


## Remarks

The  **Caption** property is a string expression that can contain up to 2,048 characters.

 If you don't set a caption for a form, button, or label, Microsoft Access will assign the object a unique name based on the object, such as "Form1".

You can use the  **Caption** property to assign an access key to a label or command button. In the caption, include an ampersand (&;) immediately preceding the character you want to use as an access key. The character will be underlined. You can press ALT plus the underlined character to move the focus to that control on a form.

Include two ampersands (&;&;) in the setting for a caption if you want to display an ampersand itself in the caption text. For example, to display "Save &; Exit", you should type  **Save &;&; Exit** in the **Caption** property box.


## See also


#### Concepts


[Page Object](page-object-access.md)

