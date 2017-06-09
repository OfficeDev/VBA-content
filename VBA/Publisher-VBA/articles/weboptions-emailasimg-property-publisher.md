---
title: WebOptions.EmailAsImg Property (Publisher)
keywords: vbapb10.chm8257545
f1_keywords:
- vbapb10.chm8257545
ms.prod: publisher
api_name:
- Publisher.WebOptions.EmailAsImg
ms.assetid: c44d3b07-2030-4901-b9df-4dcfe08c985c
ms.date: 06/08/2017
---


# WebOptions.EmailAsImg Property (Publisher)

 **True** to send the entire publication page as a single JPEG image. Read/write **Boolean**.


## Syntax

 _expression_. **EmailAsImg**

 _expression_A variable that represents an  **WebOptions** object.


### Return Value

Boolean


## Remarks

This property can increase your message's compatibility with older e-mail clients, but may result in larger file size.

This property is accessible for print publications in addition to Web publications.

The properties of the  **[WebOptions](weboptions-object-publisher.md)** object are used to specify the behavior of Web publications. This means that when any of these properties are modified, newly created Web publications will inherit the modified properties.

This property corresponds to the check box in the  **E-Mail Options** section of the **Web** tab of the **Options** dialog box.


## Example

The following example sets Microsoft Publisher to e-mail publication pages as JPEG images.


```vb
Application.WebOptions.EmailAsImg = True
```


