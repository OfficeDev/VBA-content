---
title: How Outlook Forms and Items Work Together
ms.prod: outlook
ms.assetid: e9b96721-3835-532e-990c-b11bf3affd6d
ms.date: 06/08/2017
---


# How Outlook Forms and Items Work Together

Microsoft Outlook stores its information in individual items in a folder. An Outlook item is similar to a record in a database in that it consists of a group of fields that collectively store information about the specific item.

Outlook displays the contents of an item in one of two ways: through a view in an explorer window, or through a form in an inspector window. A form usually provides a more complete display of the information and lets the user interact with the contents of the item in more ways. In a sense, a form is the principal user interface for an item. Outlook provides one or more standard forms for each type of item (mail message, contact, and so on). You can create customized versions of these forms to change the way Outlook displays items. You can display additional pages that are usually hidden, and you can add controls to those pages. Typically, these controls are bound to fields in the item so that the user can view and edit the contents of those fields. You can also customize forms by using form regions. For more information, see  [Customizing Form Pages and Form Regions](customizing-form-pages-and-form-regions.md).

Every item contains a  **Message Class** field; this field contains the name of the form that Outlook provides to view and edit the item. For example, a contact item has a default message class of "IPM.Contact". If you create a custom form called "Customer", the **Message Class** field of items using that form will contain "IPM.Contact.Customer". The message class of all Outlook items always begins with "IPM". The second part of the message class denotes the type of Outlook form that the form is based on. The third portion of the message class is present only if the form is a customized version of a standard Outlook form.

If you are creating custom forms by using form regions, the message class must be specified in the Windows Registry. For more information, see  [Specifying Form Regions in the Windows Registry](specifying-form-regions-in-the-windows-registry.md).

