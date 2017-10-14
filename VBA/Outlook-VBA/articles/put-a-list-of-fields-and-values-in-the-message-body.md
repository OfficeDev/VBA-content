---
title: Put a List of Fields and Values in the Message Body
keywords: olfm10.chm3077508
f1_keywords:
- olfm10.chm3077508
ms.prod: outlook
ms.assetid: 8e8db2cf-4918-694d-3941-8334e7aaa0cf
ms.date: 06/08/2017
---


# Put a List of Fields and Values in the Message Body

To add a list of fields and values in the body of an item, define a variable to contain the string, and then use Outlook properties that refer to the field you want to include. For example, to include the To field in the message body, use the following.


```vb
Chr(13)
```


is the return character.




```vb
MessageString = "This letter is sent to " &; Item.To &; Chr(13) 
MessageString = MessageString &; "second line goes here" 
Item.Body = MessageString
```


