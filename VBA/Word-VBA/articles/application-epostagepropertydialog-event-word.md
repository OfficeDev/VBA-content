---
title: Application.EPostagePropertyDialog Event (Word)
keywords: vbawd10.chm4000014
f1_keywords:
- vbawd10.chm4000014
ms.prod: word
api_name:
- Word.Application.EPostagePropertyDialog
ms.assetid: 6d48fb9b-7826-2897-4deb-bde202fbd05b
ms.date: 06/08/2017
---


# Application.EPostagePropertyDialog Event (Word)

Occurs when a user clicks the  **E-postage Properties** ( **Labels and Envelopes** dialog box) button or **Print Electronic Postage** button.


## Syntax

 _expression_ . **Private Sub object_EPostagePropertyDialog**( **_ByVal Doc As Document_** )

 _expression_ A variable that represents an **[Application](application-object-word.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see[Using Events with the Application Object](http://msdn.microsoft.com/library/784c4c61-7e47-3dbf-46f6-da655f786ca1%28Office.15%29.aspx).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The name of the document to which to add electronic postage.|

## Remarks

This event allows a third-party software application to intercept and show their properties dialog box.


## Example

This example displays a message when a user clicks either the  **Add Electronic Postage** button or the **Print Electronic Postage** button.


```vb
Private Sub AppWord_EPostagePropertyDialog(ByVal Doc As Document) 
 MsgBox "You have clicked a button to " &; _ 
 "display the ePostage property dialog box." 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-word.md)

