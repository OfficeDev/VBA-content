---
title: ExchangeUser.GetPicture Method (Outlook)
keywords: vbaol11.chm3485
f1_keywords:
- vbaol11.chm3485
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetPicture
ms.assetid: 4298db85-0576-4982-9592-6eae666d966a
ms.date: 06/08/2017
---


# ExchangeUser.GetPicture Method (Outlook)

Obtains an  **[IPictureDisp](http://msdn.microsoft.com/en-us/library/ms680762%28VS.85%29.aspx)** object that represents the picture of the Microsoft Exchange user that is displayed in Microsoft Outlook.


## Syntax

 _expression_ . **GetPicture**

 _expression_ A variable that represents an **[ExchangeUser](exchangeuser-object-outlook.md)** object.


### Return Value

An  **IPictureDisp** object that represents the picture of the Exchange user that is displayed in Outlook.


## Remarks

The picture of the Exchange user is stored in Active Directory and displayed in various places in Outlook, including the dialog box for  **Outlook Properties** and Contact Card.

If the picture does not exist for the user,  **GetPicture** returns **Null** ( **Nothing** for Visual Basic).

You can only call  **GetPicture** from code that runs in-process as Outlook. An **StdPicture** object cannot be marshaled across process boundaries. If you attempt to call **GetPicture** from out-of-process code, an exception occurs. For more information, see[An automation server cannot pass a pointer to the picture object's IPictureDisp implementation across process boundaries](http://support.microsoft.com/kb/150034).


## See also


#### Concepts


[ExchangeUser Object](exchangeuser-object-outlook.md)

