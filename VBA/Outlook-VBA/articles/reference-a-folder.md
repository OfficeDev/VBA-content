---
title: Reference a Folder
keywords: olfm10.chm3077416
f1_keywords:
- olfm10.chm3077416
ms.prod: outlook
ms.assetid: 65ccbabd-7ac7-ffd1-d963-e8a029152bd6
ms.date: 06/08/2017
---


# Reference a Folder

To reference a folder by the name of the folder, use the following code.


```vb
Application.GetNameSpace("MAPI").Folders("Personal Folders").Folders("Product Ideas")
```


To reference a folder by a number, use the following code. In this example, the first folder in the folder collection Personal Folders is referenced.




```vb
Application.GetNameSpace("MAPI").Folders("Personal Folders").Folders(1)
```

To reference any of the default Outlook folders, use the  **GetDefaultFolder** method. Use the appropriate constant value from the ** [OlDefaultFolders Enumeration](oldefaultfolders-enumeration-outlook.md)** to specify the folder you want to create.



```vb
Application.GetNameSpace("MAPI").GetDefaultFolder(6)
```


