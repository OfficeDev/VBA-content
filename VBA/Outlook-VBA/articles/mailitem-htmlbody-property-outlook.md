---
title: MailItem.HTMLBody Property (Outlook)
keywords: vbaol11.chm1338
f1_keywords:
- vbaol11.chm1338
ms.prod: outlook
api_name:
- Outlook.MailItem.HTMLBody
ms.assetid: c340fe05-9a99-3a32-3d6b-f2f7a568b299
ms.date: 06/08/2017
---


# MailItem.HTMLBody Property (Outlook)

Returns or sets a  **String** representing the HTML body of the specified item. Read/write.


## Syntax

 _expression_ . **HTMLBody**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

The  **HTMLBody** property should be an HTML syntax string.

Setting the  **HTMLBody** property will always update the **[Body](mailitem-body-property-outlook.md)** property immediately.


## Example

The following Visual Basic for Applications (VBA) example creates a new  **[MailItem](mailitem-object-outlook.md)** object and sets the **[BodyFormat](mailitem-bodyformat-property-outlook.md)** property to **olFormatHTML** . The body text of the e-mail item will now appear in HTML format.


```vb
Sub CreateHTMLMail() 
 
 'Creates a new e-mail item and modifies its properties. 
 
 Dim objMail As Outlook.MailItem 
 
 
 
 'Create e-mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 With objMail 
 
 'Set body format to HTML 
 
 .BodyFormat = olFormatHTML 
 
 .HTMLBody = _ 
 
 "<HTML><BODY>Enter the message text here. </BODY></HTML>" 
 
 .Display 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

