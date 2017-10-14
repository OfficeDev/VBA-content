---
title: Referencing Fields
keywords: olfm10.chm3077117
f1_keywords:
- olfm10.chm3077117
ms.prod: outlook
ms.assetid: 2ddf0f1d-f889-d631-caf2-af5d80c6b9ef
ms.date: 06/08/2017
---


# Referencing Fields

When you need to access the fields in an item, the method you use depends on whether the field is a standard built-in Outlook field, or a custom field.

In either case, you do not access the field directly. Instead, you refer to the field as a property of the item you're working with.

For example, to retrieve the text from the Subject field of a mail message, you use the  **Subject** property of the item, as shown in the following VBScript example.




```
mySubject = Item.Subject
```

If the field is a custom (user-defined) field, you access it using the  **UserProperties** property of the item, as shown in the following VBScript example. This example assumes that the item already contains a custom field named ReferredBy.



```
MyReferral = Item.UserProperties("ReferredBy")
```


