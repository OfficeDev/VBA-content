---
title: About Using VBScript in Outlook
keywords: olfm10.chm3077140
f1_keywords:
- olfm10.chm3077140
ms.prod: outlook
ms.assetid: 7c393ab9-1fa6-f615-b4ca-74b15f708809
ms.date: 06/08/2017
---


# About Using VBScript in Outlook

One way of extending form pages in a custom form is by using Microsoft Visual Basic Scripting Edition (VBScript). VBScript is a powerful scripting language based on Microsoft Visual Basic that enables you to control objects, folders, forms, items, and controls within a page of a form. For example, you can change properties and values of controls on a page, modify the default Microsoft Outlook item events, and even create automated procedures, such as mailing a notice to all the contacts in a Contacts folder.

You add VBScript code to an Outlook form to respond to  **Click** events that are fired by controls on the form page, or to respond to events fired by the items that have the same message class as the form. VBScript makes it especially easy to respond to item events because the VBScript code executes in the context of the item, so you don't have to set an object variable to point to the item. In addition, VBScript code is compact and can be contained within a form sent to other users.

With VBScript, you have full access to the Outlook object model, except for two areas: VBScript code cannot respond to events other than item and form events, and you cannot use named constants defined in the Outlook object type library.

You can also use Visual Basic for Applications in Outlook to respond to Outlook events and to create macros that automate procedures. Unlike VBScript code, however, Visual Basic for Applications code cannot be contained in a form and so cannot accompany an item that is sent to other users. Note that VBScript is only applicable to extending a form with a form page; if you are extending a form with a form region, you will not be able to use VBScript and will have to use an add-in.

