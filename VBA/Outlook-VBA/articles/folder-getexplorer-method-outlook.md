---
title: Folder.GetExplorer Method (Outlook)
keywords: vbaol11.chm1997
f1_keywords:
- vbaol11.chm1997
ms.prod: outlook
api_name:
- Outlook.Folder.GetExplorer
ms.assetid: f60bf373-802e-cb93-2152-bc6c8945edb1
ms.date: 06/08/2017
---


# Folder.GetExplorer Method (Outlook)

Returns an  **[Explorer](explorer-object-outlook.md)** object that represents a new, inactive **Explorer** object initialized with the specified folder as the current folder.


## Syntax

 _expression_ . **GetExplorer**( **_DisplayMode_** )

 _expression_ A variable that represents a **Folder** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _DisplayMode_|Optional| **Variant**|The display mode of the folder. Can be one of the constants in the  **[OlFolderDisplayMode](olfolderdisplaymode-enumeration-outlook.md)** enumeration.|

### Return Value

An  **Explorer** object that represents a new, inactive Explorer initialized with the specified folder as the current folder.


## Remarks

This method is useful for returning a new  **Explorer** object in which to display the folder, as opposed to using the **[ActiveExplorer](application-activeexplorer-method-outlook.md)** method and setting the **[CurrentFolder](explorer-currentfolder-property-outlook.md)** property.

The  **[Explorer.Display](explorer-display-method-outlook.md)** method can be used to activate or show the **Explorer** .

The  **GetExplorer** method takes an optional argument of an **OlFolderDisplayMode** constant.

By default, the new  **Explorer** will be displayed in the Normal mode ( **olFolderDisplayNormal** ) with all interface elements displayed: a message panel on the right and the Navigation Pane on the left. The exception to this rule is when you are calling **GetExplorer** on delegated folders that are in No-Navigation mode ( **olFolderDisplayNoNavigation** ) by default. You can apply more restrictions to a default mode, but you cannot lessen the restrictions by changing the **OlFolderDisplayMode** .

The explorer can also be displayed in Folder-Only mode ( **olFolderDisplayFolderOnly** ). This mode is essentially the same as the Normal mode ( **olFolderDisplayNormal** ) in that it too displays the Navigation Pane on the left.

 The most restrictive mode you can use is No-Navigation mode ( **olFolderDisplayNoNavigation** ). In this mode, the **Explorer** will display with no folder list, no drop-down folder list, and any "Go"-type menu/command bar options should be disabled. Basically, the user should not be able to navigate to any other folder within that **Explorer** window. By default, a delegated (shared) folder appears in No-Navigation mode.


## Example

This Visual Basic for Applications (VBA) example uses the  **GetExplorer** method to return a new, inactive explorer for the default Contacts folder and then displays the explorer in the default mode of **olFolderDisplayNormal** .


```vb
Sub ActivateContactExplorer() 
 
 Dim nsp As Outlook.NameSpace 
 
 Dim mpfContacts As Outlook.Folder 
 
 Dim expContacts As Outlook.Explorer 
 
 
 
 Set nsp = Application.GetNamespace("MAPI") 
 
 Set mpfContacts = nsp.GetDefaultFolder(olFolderContacts) 
 
 Set expContacts = mpfContacts.GetExplorer 
 
 expContacts.Activate 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

