---
title: Application.CreateItem Method (Outlook)
keywords: vbaol11.chm714
f1_keywords:
- vbaol11.chm714
ms.prod: outlook
api_name:
- Outlook.Application.CreateItem
ms.assetid: e5fbf367-db16-5042-823e-68e6b805e612
ms.date: 06/08/2017
---


# Application.CreateItem Method (Outlook)

Creates and returns a new Microsoft Outlook item.


## Syntax

 _expression_ . **CreateItem**( **_ItemType_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ItemType_|Required| **[OlItemType](olitemtype-enumeration-outlook.md)**|The Outlook item type for the new item.|

### Return Value

An  **Object** value that represents the new Outlook item.


## Remarks

The  **CreateItem** method can only create default Outlook items. To create new items using a custom form, use the **[Add](items-add-method-outlook.md)** method on the **[Items](items-object-outlook.md)** collection.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new  **[MailItem](mailitem-object-outlook.md)** object and sets the **BodyFormat** property to **olFormatHTML** . The Body text of the e-mail item will now appear in HTML format.


```vb
Sub CreateHTMLMail() 
 
 'Creates a new e-mail item and modifies its properties 
 
 Dim objMail As Outlook.MailItem 
 
 
 
 'Create e-mail item 
 
 Set objMail = Application.CreateItem(olMailItem) 
 
 With objMail 
 
 'Set body format to HTML 
 
 .BodyFormat = olFormatHTML 
 
 .HTMLBody = "<HTML><H2>The body of this message will appear in HTML.</H2><BODY> Please enter the message text here. </BODY></HTML>" 
 
 .Display 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)
#### Other resources



[How to: Import Appointment XML Data into Outlook Appointment Objects](http://msdn.microsoft.com/library/ecfd3849-877b-01ad-2b76-1a54e980f6e2%28Office.15%29.aspx)

