---
title: DefaultWebOptions.SaveNewWebPagesAsWebArchives Property (Word)
keywords: vbawd10.chm165871634
f1_keywords:
- vbawd10.chm165871634
ms.prod: word
api_name:
- Word.DefaultWebOptions.SaveNewWebPagesAsWebArchives
ms.assetid: a2c8a225-431e-9292-d081-bd71d27aae9c
ms.date: 06/08/2017
---


# DefaultWebOptions.SaveNewWebPagesAsWebArchives Property (Word)

 **True** for Microsoft Word to save new Web pages in the Single File Web Page (formerly known as Web Archive) format. Read/write **Boolean** .


## Syntax

 _expression_ . **SaveNewWebPagesAsWebArchives**

 _expression_ An expression that returns a **[DefaultWebOptions](defaultweboptions-object-word.md)** object.


## Remarks

Setting the  **SaveNewWebPagesAsWebArchives** property won't change the format of any saved Web pages. To change their format, you must individually open them and then use the **[SaveAs2](document-saveas2-method-word.md)** method to set the Web page format.


## Example

This example enables the  **SaveNewWebPagesAsWebArchives** property so that when Web pages are saved, they are saved in the Single File Web Page format.


```vb
Sub SetWebOption() 
 Application.DefaultWebOptions _ 
 .SaveNewWebPagesAsWebArchives = True 
End Sub
```


## See also


#### Concepts


[DefaultWebOptions Object](defaultweboptions-object-word.md)

