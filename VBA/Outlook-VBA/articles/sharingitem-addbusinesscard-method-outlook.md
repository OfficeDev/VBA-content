---
title: SharingItem.AddBusinessCard Method (Outlook)
keywords: vbaol11.chm3217
f1_keywords:
- vbaol11.chm3217
ms.prod: outlook
api_name:
- Outlook.SharingItem.AddBusinessCard
ms.assetid: fa3fa071-b43c-c2d1-7d7c-dc52ab9a1681
ms.date: 06/08/2017
---


# SharingItem.AddBusinessCard Method (Outlook)

Appends contact information based on the Electronic Business Card (EBC) associated with the specified  **[ContactItem](contactitem-object-outlook.md)** object to the **[SharingItem](sharingitem-object-outlook.md)** object.


## Syntax

 _expression_ . **AddBusinessCard**( **_contact_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _contact_|Required| **ContactItem**|The contact item from which to obtain the business card information.|

## Remarks

This method adds contact information, generated from the information stored in the  **ContactItem** object, to the existing **SharingItem** object. The information included depends on the value of the **[BodyFormat](sharingitem-bodyformat-property-outlook.md)** property for the **SharingItem** object:



| **Property value**| **Result**|
| **olFormatPlain**|A vCard (.vcf) file is created and added to the  **[Attachments](attachments-object-outlook.md)** collection of the **SharingItem** object.|
| **olFormatRichText**|A vCard (.vcf) file is created and added to the  **Attachments** collection of the **SharingItem** object.|
| **olFormatHTML**|An image of the business card is generated and included in the  **[Body](mailitem-body-property-outlook.md)** property of the **SharingItem** object, and a vCard (.vcf) file is created and added to the **[Attachments](attachments-object-outlook.md)** collection of the **SharingItem** object.|

 **Note**  The attached vCard file contains only the contact information included in the Electronic Business Card associated with the  **ContactItem** object. Any contact information not displayed in the Electronic Business Card is excluded from the vCard file.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

