---
title: ContactItem.ForwardAsBusinessCard Method (Outlook)
keywords: vbaol11.chm1094
f1_keywords:
- vbaol11.chm1094
ms.prod: outlook
api_name:
- Outlook.ContactItem.ForwardAsBusinessCard
ms.assetid: 2f1a74c3-86f0-a054-75e2-272dbb261fb7
ms.date: 06/08/2017
---


# ContactItem.ForwardAsBusinessCard Method (Outlook)

Creates a new  **[MailItem](mailitem-object-outlook.md)** object containing contact information and, optionally, an Electronic Business Card (EBC) image based on the specified **[ContactItem](contactitem-object-outlook.md)** object.


## Syntax

 _expression_ . **ForwardAsBusinessCard**

 _expression_ An expression that returns a **ContactItem** object.


### Return Value

A  **MailItem** object that represents the new e-mail item containing the business card information.


## Remarks

This method creates a new Outlook mail item based on the information stored in the  **ContactItem** object. The information included in the Outlook mail item depends on the value of the **[BodyFormat](mailitem-bodyformat-property-outlook.md)** property for the **MailItem** object:



| **Property value**| **Result**|
| **olFormatPlain**|A vCard (.vcf) file is created and added to the  **[Attachments](attachments-object-outlook.md)** collection of the **MailItem** object.|
| **olFormatRichText**|A vCard file is created and added to the  **Attachments** collection of the **MailItem** object.|
| **olFormatHTML**|An image of the Electronic Business Card is generated and included in the  **[Body](mailitem-body-property-outlook.md)** property of the **MailItem** object, and a vCard file is created and added to the **[Attachments](attachments-object-outlook.md)** collection of the **MailItem** object.|

 **Note**  The attached vCard file contains only the contact information included in the Electronic Business Card associated with the  **ContactItem** object. Any contact information not displayed in the Electronic Business Card is excluded from the vCard file.


## See also


#### Concepts


[ContactItem Object](contactitem-object-outlook.md)

