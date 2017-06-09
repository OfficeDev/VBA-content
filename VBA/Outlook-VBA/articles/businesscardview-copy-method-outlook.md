---
title: BusinessCardView.Copy Method (Outlook)
keywords: vbaol11.chm2922
f1_keywords:
- vbaol11.chm2922
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Copy
ms.assetid: 9a0a1a14-87bd-ff53-6643-5e11a07733a1
ms.date: 06/08/2017
---


# BusinessCardView.Copy Method (Outlook)

Creates a new  **[View](view-object-outlook.md)** object based on the existing **[BusinessCardView](businesscardview-object-outlook.md)** object.


## Syntax

 _expression_ . **Copy**( **_Name_** , **_SaveOption_** )

 _expression_ An expression that returns a **BusinessCardView** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)**|The save option for the new view.|

### Return Value

A  **View** object that represents the new view.


## Example

The following Visual Basic for Applications (VBA) example creates a copy of a  **BusinessCardView** object, named "New Card View", and saves it in the **Contacts** default folder. To run this example, you need to first create a **BusinessCardView** object named "Card View" either programmatically or by using the Microsoft Outlook user interface.


```vb
Sub CopyBusinessCardView() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objNewView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Copy the existing view. 
 
 Set objNewView = objViews("Card View").Copy( _ 
 
 "New Card View", _ 
 
 olViewSaveOptionThisFolderEveryone) 
 
 
 
End Sub
```


## See also


#### Concepts


[BusinessCardView Object](businesscardview-object-outlook.md)

