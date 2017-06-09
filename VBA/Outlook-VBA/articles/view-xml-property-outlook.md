---
title: View.XML Property (Outlook)
keywords: vbaol11.chm2495
f1_keywords:
- vbaol11.chm2495
ms.prod: outlook
api_name:
- Outlook.View.XML
ms.assetid: a933daaa-370f-2ed3-0a59-86f766a1f2c8
ms.date: 06/08/2017
---


# View.XML Property (Outlook)

Returns or sets a  **String** value that specifies the XML definition of the current view. Read/write.


## Syntax

 _expression_ . **XML**

 _expression_ A variable that represents a **View** object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the XML property for that view.

To programmatically add a custom field to a view, use the  **[Add](viewfields-add-method-outlook.md)** method of the **[ViewFields](viewfields-object-outlook.md)** object. This is the recommended way to dynamically change the view over setting the **[XML](view-xml-property-outlook.md)** property of the **[View](view-object-outlook.md)** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates an instance of the  **[Views](views-object-outlook.md)** collection and displays the XML definition of a view called "Table View". If the view does not exist, it creates one.


```vb
Sub DisplayViewDef() 
 
 'Displays the XML definition of a View object 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Return a view called Table View if it already exists, else create one 
 
 Set objView = objViews.Item("Table View") 
 
 If objView Is Nothing Then 
 
 Set objView = objViews.Add("Table View", olTableView, olViewSaveOptionAllFoldersOfType) 
 
 End If 
 
 MsgBox objView.XML 
 
End Sub
```

The following are the modified properties that are visible in the following XML source code. In addition to the property definitions, the XML source also defines any objects that make up the view. The following example displays the XML definition of columns that appear in the above view.




```XML
<heading>Flag Status</heading>     <prop>http://schemas.microsoft.com/mapi/proptag/0x10900003</prop>     <type>i4</type>     <bitmap>1</bitmap>     <style>padding-left:3px;text-align:center;padding-left:3px</style> </column> <column>     <format>boolicon</format>     <heading>Attachment</heading>     <prop>urn:schemas:httpmail:hasattachment</prop>     <type>boolean</type>     <bitmap>1</bitmap>     <style>padding-left:3px;text-align:center;padding-left:3px</style>     <displayformat>3</displayformat> </column>
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

