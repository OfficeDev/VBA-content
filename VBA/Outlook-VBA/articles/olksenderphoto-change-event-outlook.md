---
title: OlkSenderPhoto.Change Event (Outlook)
keywords: vbaol11.chm1000492
f1_keywords:
- vbaol11.chm1000492
ms.prod: outlook
api_name:
- Outlook.OlkSenderPhoto.Change
ms.assetid: a4d58172-a16f-6084-9230-af2c3cefa552
ms.date: 06/08/2017
---


# OlkSenderPhoto.Change Event (Outlook)

Occurs when the sender's contact picture has changed. 


## Syntax

 _expression_ . **Change**

 _expression_ A variable that represents an **OlkSenderPhoto** object.


## Remarks

The change of the sender's contact picture usually means that the  **[PreferredWidth](olksenderphoto-preferredwidth-property-outlook.md)** and **[PreferredHeight](olksenderphoto-preferredheight-property-outlook.md)** properties have changed as well. Therefore, this event can be used as an indication of the necessity to resize the control.


## See also


#### Concepts


[OlkSenderPhoto Object](olksenderphoto-object-outlook.md)

