---
title: MailItem.BodyFormat Property (Outlook)
keywords: vbaol11.chm1372
f1_keywords:
- vbaol11.chm1372
ms.prod: outlook
api_name:
- Outlook.MailItem.BodyFormat
ms.assetid: f635a0bc-20b7-206c-f558-a4ca2519670f
ms.date: 06/08/2017
---


# MailItem.BodyFormat Property (Outlook)

Returns or sets an  **[OlBodyFormat](olbodyformat-enumeration-outlook.md)** constant indicating the format of the body text. Read/write.


## Syntax

 _expression_ . **BodyFormat**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The body text format determines the standard used to display the text of the message. Microsoft Outlook provides three body text format options: Plain Text, Rich Text (RTF), and HTML.

All text formatting will be lost when the  **BodyFormat** property is switched from RTF to HTML and vice-versa.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new  **[MailItem](mailitem-object-outlook.md)** object and sets the **BodyFormat** property to **olFormatHTML** . The body text of the e-mail item will now appear in HTML format.


```vb
Sub CreateHTMLMail() 
 
 'Creates a new e-mail item and modifies its properties. 
 
 Dim objMail As MailItem 
 
 
 
 'Create mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 With objMail 
 
 'Set body format to HTML 
 
 .BodyFormat = olFormatHTML 
 
 .HTMLBody = "<HTML><H2>The body of this message will appear in HTML.</H2><BODY>Type the message text here. </BODY></HTML>" 
 
 .Display 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

