---
title: AccountSelector.SelectedAccount Property (Outlook)
keywords: vbaol11.chm3453
f1_keywords:
- vbaol11.chm3453
ms.prod: outlook
api_name:
- Outlook.AccountSelector.SelectedAccount
ms.assetid: ecb0990b-16d6-51fb-bfc9-038b8dcca383
ms.date: 06/08/2017
---


# AccountSelector.SelectedAccount Property (Outlook)

Returns an  **[Account](account-object-outlook.md)** object that represents the selected account in the Microsoft Office Backstage view in Microsoft Outlook. Read-only.


## Syntax

 _expression_ . **SelectedAccount**

 _expression_ A variable that represents an **[AccountSelector](accountselector-object-outlook.md)** object.


## Remarks

In the Outlook user interface, you can select an account explicitly on the  **Info** tab of the Backstage view, or you can select an account implicitly when you select a folder in a list of folders. The **SelectedAccount** property represents the currently selected account in the Backstage view for a given instance of the **[Explorer](explorer-object-outlook.md)** object. To determine the account that is selected implicitly, identify the **Account** object that has its **[DefaultStore](namespace-defaultstore-property-outlook.md)** property equal to the **[Store](folder-store-property-outlook.md)** property of the current folder (which is represented by **[Explorer.CurrentFolder](explorer-currentfolder-property-outlook.md)** ).

This property returns  **Null** ( **Nothing** in Visual Basic) if no accounts are defined in the session's profile; that is, the **[Namespace.Accounts.Count](accounts-count-property-outlook.md)** property is 0.


## See also


#### Concepts


[AccountSelector Object](accountselector-object-outlook.md)

