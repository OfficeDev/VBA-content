---
title: MailItem.GetInspector Property (Outlook)
keywords: vbaol11.chm1305
f1_keywords:
- vbaol11.chm1305
ms.prod: outlook
api_name:
- Outlook.MailItem.GetInspector
ms.assetid: 9ba8bdbf-1dd5-eaff-3889-33433e3cb3fa
ms.date: 06/08/2017
---


# MailItem.GetInspector Property (Outlook)

Returns an  **[Inspector](inspector-object-outlook.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

 _expression_ . **GetInspector**

 _expression_ A variable that represents a **MailItem** object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](application-activeinspector-method-outlook.md)** method and setting the **[Inspector.CurrentItem](inspector-currentitem-property-outlook.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## Example

This Visual Basic for Applications (VBA) example shows a function  `InsertBodyTextInWordEditor` that creates a mail item, assigns it a title and adds text for the body. The function sets the **[Subject](mailitem-subject-property-outlook.md)** property to assign the title "Testing...". It then calls the **[Display](mailitem-display-method-outlook.md)** method to open the mail item in an inspector. To insert text in a Word editor as the body of the mail item, the function uses the **[Document](http://msdn.microsoft.com/library/8d83487a-2345-a036-a916-971c9db5b7fb%28Office.15%29.aspx)** object and **[Range](http://msdn.microsoft.com/library/15a7a1c4-5f3f-5b6e-60e9-29688de3f274%28Office.15%29.aspx)** object in the Word object model. The function uses the item's **GetInspector** property to get the existing **Inspector** object, and then uses the **[Inspector.WordEditor](inspector-wordeditor-property-outlook.md)** property to obtain a **Word.Document** object for the item. Using the **Word.Document** object, the function accesses the **Word.Range** object and inserts text into the body of the item.

Since this example accesses the Word object model, you must first add a reference to the Microsoft Word Object Library to compile the example successfully.




```vb
Sub InsertBodyTextInWordEditor() 
 Dim myItem As Outlook.MailItem 
 Dim myInspector As Outlook.Inspector 
 'You must add a reference to the Microsoft Word Object Library 
 'before this sample will compile 
 Dim wdDoc As Word.Document 
 Dim wdRange As Word.Range 
 
 On Error Resume Next 
 Set myItem = Application.CreateItem(olMailItem) 
 myItem.Subject = "Testing..." 
 myItem.Display 
 'GetInspector property returns Inspector 
 Set myInspector = myItem.GetInspector 
 'Obtain the Word.Document for the Inspector 
 Set wdDoc = myInspector.WordEditor 
 If Not (wdDoc Is Nothing) Then 
 'Use the Range object to insert text 
 Set wdRange = wdDoc.Range(0, wdDoc.Characters.Count) 
 wdRange.InsertAfter ("Hello world!") 
 End If 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

