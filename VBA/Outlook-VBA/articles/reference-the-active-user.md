---
title: Reference the Active User
keywords: olfm10.chm3077415
f1_keywords:
- olfm10.chm3077415
ms.prod: outlook
ms.assetid: dc8e5e24-51a0-1b16-649e-6b24e0fa9b56
ms.date: 06/08/2017
---


# Reference the Active User

Use  **[Application.GetNamespace](application-getnamespace-method-outlook.md)** to return the Outlook **[NameSpace](namespace-object-outlook.md)** object from the **[Application](application-object-outlook.md)** object, and then use the **[NameSpace.CurrentUser](namespace-currentuser-property-outlook.md)** property to return a **[Recipient](recipient-object-outlook.md)** object repesenting the active user, as shown in the following example.


```vb
Set myUser = Application.GetNameSpace("MAPI").CurrentUser
```


