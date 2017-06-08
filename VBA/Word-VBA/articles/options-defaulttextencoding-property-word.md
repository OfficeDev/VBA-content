---
title: Options.DefaultTextEncoding Property (Word)
keywords: vbawd10.chm162988475
f1_keywords:
- vbawd10.chm162988475
ms.prod: word
api_name:
- Word.Options.DefaultTextEncoding
ms.assetid: 068f0ddd-efb4-9bb3-4544-79d390e87f59
ms.date: 06/08/2017
---


# Options.DefaultTextEncoding Property (Word)

Returns or sets an  **MsoEncoding** constant representing the code page, or character set, that Microsoft Word uses for all documents saved as encoded text files. Read/write.


## Syntax

 _expression_ . **DefaultTextEncoding**

 _expression_ Required. A variable that represents an **[Options](options-object-word.md)** collection.


## Remarks

Use the  **TextEncoding** property to set the encoding for an individual document. To set encoding for HTML documents, use the **Encoding** property.


## Example

This example sets the global text encoding to the Western code page. This means that Word will save all encoded text files using the Western code page.


```vb
Sub DefaultEncode() 
 Application.Options.DefaultTextEncoding = msoEncodingWestern 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

