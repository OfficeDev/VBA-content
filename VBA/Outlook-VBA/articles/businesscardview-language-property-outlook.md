---
title: BusinessCardView.Language Property (Outlook)
keywords: vbaol11.chm2926
f1_keywords:
- vbaol11.chm2926
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Language
ms.assetid: 4ddc6c63-3402-15ae-bcb7-7eab0d423bb3
ms.date: 06/08/2017
---


# BusinessCardView.Language Property (Outlook)

Returns or sets a  **String** value that represents the language setting for the object that defines the language used in the menu. Read/write.


## Syntax

 _expression_ . **Language**

 _expression_ A variable that represents a **BusinessCardView** object.


## Remarks

The  **Language** property uses a **String** to represent an ISO language tag. For example, the string "EN-US" represents the ISO code for "United States - English."

If a valid language code is specified, the object will only be available in the  **View** menu for the specified language type. If no value is specified, the object item is available for all language types. The default value for this property is an empty string.


## Example

The following Visual Basic for Applications (VBA) example sets the language type of all  **[View](view-object-outlook.md)** objects of type **olBusinessCArdView** to U.S. English.


```vb
Sub SetLanguage() 
 
 'Sets the language of all table views to U.S. English. 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objViews = _ 
 
 Application.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Views 
 
 'Iterate through each view in the collection. 
 
 For Each objView In objViews 
 
 Debug.Print objView.Name 
 
 'If view is of type olBusinessCardVIew then set language. 
 
 If objView.ViewType = olBusinessCardView And objView.Standard = False Then 
 
 objView.Language = "EN-US" 
 
 End If 
 
 Next objView 
 
End Sub
```


## See also


#### Concepts


[BusinessCardView Object](businesscardview-object-outlook.md)

