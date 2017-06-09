---
title: MailItem.AddBusinessCard Method (Outlook)
keywords: vbaol11.chm1389
f1_keywords:
- vbaol11.chm1389
ms.prod: outlook
api_name:
- Outlook.MailItem.AddBusinessCard
ms.assetid: a30d201b-3073-11c1-0f0c-81c7a3aba6e2
ms.date: 06/08/2017
---


# MailItem.AddBusinessCard Method (Outlook)

Appends contact information based on the Electronic Business Card (EBC) associated with the specified  **[ContactItem](contactitem-object-outlook.md)** object to the **[MailItem](mailitem-object-outlook.md)** object.


## Syntax

 _expression_ . **AddBusinessCard**( **_contact_** )

 _expression_ An expression that returns a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _contact_|Required| **ContactItem**|The contact item from which to obtain the business card information.|

## Remarks

This method adds contact information, generated from the information stored in the  **ContactItem** object, to the existing **MailItem** object. The information included depends on the value of the **[BodyFormat](mailitem-bodyformat-property-outlook.md)** property for the **MailItem** object:



| **Property value**| **Result**|
| **olFormatPlain**|A vCard (.vcf) file is created and added to the  **[Attachments](attachments-object-outlook.md)** collection of the **MailItem** object.|
| **olFormatRichText**|A vCard (.vcf) file is created and added to the  **Attachments** collection of the **MailItem** object.|
| **olFormatHTML**|An image of the business card is generated and included in the  **[Body](mailitem-body-property-outlook.md)** property of the **MailItem** object, and a vCard (.vcf) file is created and added to the **[Attachments](attachments-object-outlook.md)** collection of the **MailItem** object.|

 **Note**  The attached vCard file contains only the contact information included in the Electronic Business Card associated with the  **ContactItem** object. Any contact information not displayed in the Electronic Business Card is excluded from the vCard file.


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

