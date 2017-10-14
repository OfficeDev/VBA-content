---
title: SelectNamesDialog.Session Property (Outlook)
keywords: vbaol11.chm823
f1_keywords:
- vbaol11.chm823
ms.prod: outlook
api_name:
- Outlook.SelectNamesDialog.Session
ms.assetid: 99f445e8-190b-fa26-319f-ff7783b27795
ms.date: 06/08/2017
---


# SelectNamesDialog.Session Property (Outlook)

Returns the  **[NameSpace](namespace-object-outlook.md)** object for the current session. Read-only.


## Syntax

 _expression_ . **Session**

 _expression_ A variable that represents a **SelectNamesDialog** object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](application-getnamespace-method-outlook.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


#### Concepts


[SelectNamesDialog Object](selectnamesdialog-object-outlook.md)

