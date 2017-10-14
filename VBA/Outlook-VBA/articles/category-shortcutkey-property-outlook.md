---
title: Category.ShortcutKey Property (Outlook)
keywords: vbaol11.chm2428
f1_keywords:
- vbaol11.chm2428
ms.prod: outlook
api_name:
- Outlook.Category.ShortcutKey
ms.assetid: c78f882a-ab02-5218-e71f-362c86b4dfe1
ms.date: 06/08/2017
---


# Category.ShortcutKey Property (Outlook)

Returns or sets an  **[OlCategoryShortcutKey](olcategoryshortcutkey-enumeration-outlook.md)** constant that specifies the shortcut key used by the **[Category](category-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **ShortcutKey**

 _expression_ A variable that represents a **Category** object.


## Remarks

Any  **OlCategoryShortcutKey** constant other than **olCategoryShortcutKeyNone** can only be used by one **Category** object at any given time. Setting the value of this property to an **OlCategoryShortcutKey** constant already in use sets the **ShortcutKey** property of the **Category** object already using the specified value to **olCategoryShortcutKeyNone** .


## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing shortcut key assignments for each  **Category** object contained in the **[Categories](namespace-categories-property-outlook.md)** collection associated with the default **[NameSpace](namespace-object-outlook.md)** object.


```vb
Private Sub ListShortcutKeys() 
 
 Dim objNameSpace As NameSpace 
 
 Dim objCategory As Category 
 
 Dim strOutput As String 
 
 
 
 ' Obtain a NameSpace object reference. 
 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 
 
 ' Check if the Categories collection for the Namespace 
 
 ' contains one or more Category objects. 
 
 If objNameSpace.Categories.Count > 0 Then 
 
 
 
 ' Enumerate the Categories collection, checking 
 
 ' the value of the ShortcutKey property for 
 
 ' each Category object. 
 
 For Each objCategory In objNameSpace.Categories 
 
 
 
 ' Add the name of the Category object to 
 
 ' the output string. 
 
 strOutput = strOutput &; objCategory.Name 
 
 
 
 ' Add information about the assigned shortcut key 
 
 ' to the output string. 
 
 Select Case objCategory.ShortcutKey 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyNone 
 
 strOutput = strOutput &; ": No shortcut key" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF2 
 
 strOutput = strOutput &; ": Ctrl+F2" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF3 
 
 strOutput = strOutput &; ": Ctrl+F3" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF4 
 
 strOutput = strOutput &; ": Ctrl+F4" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF5 
 
 strOutput = strOutput &; ": Ctrl+F5" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF6 
 
 strOutput = strOutput &; ": Ctrl+F6" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF7 
 
 strOutput = strOutput &; ": Ctrl+F7" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF8 
 
 strOutput = strOutput &; ": Ctrl+F8" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF9 
 
 strOutput = strOutput &; ": Ctrl+F9" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF10 
 
 strOutput = strOutput &; ": Ctrl+F10" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF11 
 
 strOutput = strOutput &; ": Ctrl+F11" &; vbCrLf 
 
 Case OlCategoryShortcutKey.olCategoryShortcutKeyCtrlF12 
 
 strOutput = strOutput &; ": Ctrl+F12" &; vbCrLf 
 
 Case Else 
 
 strOutput = strOutput &; ": Unknown" &; vbCrLf 
 
 End Select 
 
 Next 
 
 End If 
 
 
 
 ' Display the output string. 
 
 MsgBox strOutput 
 
 
 
 ' Clean up. 
 
 Set objCategory = Nothing 
 
 Set objNameSpace = Nothing 
 
 
 
End Sub
```


## See also


#### Concepts


[Category Object](category-object-outlook.md)

