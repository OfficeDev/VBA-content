---
title: NameSpace.Offline Property (Outlook)
keywords: vbaol11.chm773
f1_keywords:
- vbaol11.chm773
ms.prod: outlook
api_name:
- Outlook.NameSpace.Offline
ms.assetid: c62112d5-e50f-bd6a-bb3b-7c1818752d8b
ms.date: 06/08/2017
---


# NameSpace.Offline Property (Outlook)

Returns a  **Boolean** indicating **True** if Outlook is offline (not connected to an Exchange server), and **False** if online (connected to an Exchange server). Read-only.


## Syntax

 _expression_ . **Offline**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

The Offline property returns valid information only for an Exchange profile. It is not intended for non-Exchange account types such as POP3, IMAPI, and HTTP.

If the  **[NameSpace.ExchangeConnectionMode](namespace-exchangeconnectionmode-property-outlook.md)** property is **olOffline** or **olDisconnected** , the **Offline** property will return **True** . If the **ExchangeConnectionMode** property is **olOnline** , **olConnected** , or **olConnectedHeaders** , the **Offline** property will return **False** .


## Example

The following Microsoft Visual Basic for Applications (VBA) example returns  **True** or **False** depending on whether the **NameSpace** object is currently online.


```vb
Sub Off() 
 
 'Determines whether Outlook is currently offline. 
 
 Dim nmsName As Outlook.NameSpace 
 
 
 
 Set nmsName = Application.GetNamespace("MAPI") 
 
 MsgBox nmsName.Offline 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

