---
title: Customize and Share Business Cards
ms.prod: outlook
ms.assetid: d29fd962-ea5f-040d-e9af-e8ab70595832
ms.date: 06/08/2017
---


# Customize and Share Business Cards

Contact information stored in Microsoft Outlook can be represented as an Electronic Business Card (EBC), in which the layout and formatting of the information contained in a  **[ContactItem](contactitem-object-outlook.md)** object can be customized for that contact item. An Electronic Business Card can be shared with other users and can be used as a signature on Outlook mail items.

 The **ContactItem** object has a default business card design associated with it at the time the object is created, and this design can be changed at any time either programmatically or by using the **Edit Business Card** dialog box. Only one Electronic Business Card design can be defined for a single **ContactItem** object. You can use the ShowBusinessCardEditor method of the ContactItem object to programmatically display the Edit Business Card dialog box. For more information about creating an Electronic Business Card design for a **ContactItem** object using the **Edit Business Card** dialog box, search for the topic "Create Electronic Business Cards" in the Outlook Help.

Several methods are provided in Office Outlook 2007 to share contact information, including Electronic Business Cards. You can use the  **[ForwardAsVcard](contactitem-forwardasvcard-method-outlook.md)** and **[ForwardAsBusinessCard](contactitem-forwardasbusinesscard-method-outlook.md)** method of the **ContactItem** object in to create a new **[MailItem](mailitem-object-outlook.md)** object that contains the contact information from the specified **ContactItem** attached as a vCard (.vcf) file, or you can use the **[AddBusinessCard](mailitem-addbusinesscard-method-outlook.md)** method of the **MailItem** object to attach the contact information for a specified **ContactItem** as a vCard file.

If you use the  **ForwardAsBusinessCard** or **AddBusinessCard** methods, the Electronic Business Card is also appended to the body of the mail item if the **[BodyFormat](mailitem-bodyformat-property-outlook.md)** property of the **MailItem** object is set to **olFormatHTML**. You can also use the  **[SaveBusinessCardImage](contactitem-savebusinesscardimage-method-outlook.md)** method of the **ContactItem** object to save an Electronic Business Card as a Portable Network Graphics (.png) image file.

