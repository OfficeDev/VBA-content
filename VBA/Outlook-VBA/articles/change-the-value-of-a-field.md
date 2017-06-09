---
title: Change the Value of a Field
keywords: olfm10.chm3077356
f1_keywords:
- olfm10.chm3077356
ms.prod: outlook
ms.assetid: a788cb9c-e2bb-b4f9-78f9-b7244ee18431
ms.date: 06/08/2017
---


# Change the Value of a Field

To change the value of an Outlook field, use the property name of the associated standard field. For example, to change the value of the Subject field, use the following code.


```
Item.Subject = "New Subject"
```


To change the value of a custom field, use the following code to refer to a custom field.




```
Item.UserProperties.Find("MyProperty").Value = "New Value"
```


