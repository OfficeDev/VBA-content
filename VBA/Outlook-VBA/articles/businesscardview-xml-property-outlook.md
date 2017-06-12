---
title: BusinessCardView.XML Property (Outlook)
keywords: vbaol11.chm2932
f1_keywords:
- vbaol11.chm2932
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.XML
ms.assetid: da381070-28e8-bace-b15f-1c01a35491b8
ms.date: 06/08/2017
---


# BusinessCardView.XML Property (Outlook)

Returns or sets a  **String** value that specifies the XML definition of the current view. Read/write.


## Syntax

 _expression_ . **XML**

 _expression_ An expression that returns a **BusinessCardView** object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, create a view by using the Outlook user interface and then retrieve the  **XML** property for that view.


## Example

The following Visual Basic for Applications (VBA) example enumerates the  **[Views](views-object-outlook.md)** collection of the **Contacts** default folder and displays the XML definition of a **[BusinessCardView](businesscardview-object-outlook.md)** object named "Card View".


```vb
Sub DisplayBusinessCardViewDef() 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Return a view called "Card View." If the view 
 
 ' doesn't already exist, create it. 
 
 Set objView = objViews.Item("Card View") 
 
 If objView Is Nothing Then 
 
 Set objView = objViews.Add( _ 
 
 "Card View", _ 
 
 olBusinessCardView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 End If 
 
 
 
 ' Display the XML definition for the view. 
 
 ' Note that the definition may be truncated 
 
 ' due to the display limitations of the 
 
 ' MsgBox function. 
 
 MsgBox objView.XML 
 
End Sub
```


## See also


#### Concepts


[BusinessCardView Object](businesscardview-object-outlook.md)

