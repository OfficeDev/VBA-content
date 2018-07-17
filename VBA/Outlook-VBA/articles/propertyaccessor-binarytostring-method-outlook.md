---
title: PropertyAccessor.BinaryToString Method (Outlook)
keywords: vbaol11.chm1977
f1_keywords:
- vbaol11.chm1977
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.BinaryToString
ms.assetid: 4a3801af-0a7c-4b8a-7367-600c09047b28
ms.date: 06/08/2017
---


# PropertyAccessor.BinaryToString Method (Outlook)

Converts the array of bytes specified by  _Value_ to a **String** .


## Syntax

 _expression_ . **BinaryToString**( **_Value_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **Variant**|Represents the array of bytes to be converted.|

### Return Value

 A hexadecimal **String** that represents the converted value.


## Remarks

For more information on type conversion when using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## Example

 The Outlook object model exposes an **EntryID** property for item objects to obtain the Entry ID of an item. This property is a string representing the value of the MAPI property, **PR_ENTRYID** , of that item. Aside from the **EntryID** property, you can also use the **[PropertyAccessor.GetProperty](propertyaccessor-getproperty-method-outlook.md)** method to obtain the value of **PR_ENTRYID** for an item, and use **PropertyAccessor.BinaryToString** to convert that value to a string. This string should match the **EntryID** property value for the same item. The following code sample shows the equivalence of the Entry ID returned by the **PropertyAccessor.GetProperty** method and the Entry ID returned by the **EntryID** property for each item in the Inbox.


```vb
Sub TestEntryIDs() 
 Dim oMsg As Object 
 Dim oFolder As Outlook.Folder 
 Dim oItems As Outlook.Items 
 Dim oPA As Outlook.PropertyAccessor 
 Dim EntryID1 As String, EntryID2 As String, EntryIDProperty As String 
 
 'This is the MAPI property PR_ENTRYID referenced with its MAPI proptag namespace 
 EntryIDProperty = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102" 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 Set oItems = oFolder.Items 
 For Each oMsg In oItems 
 Set oPA = oMsg.PropertyAccessor 
 'First use the EntryID property of the item 
 EntryID1 = oMsg.EntryID 
 'Then use the PropertyAccessor 
 EntryID2 = oPA.BinaryToString(oPA.GetProperty(EntryIDProperty)) 
 'The string equivalents of the two Entry IDs should be the same 
 If EntryID1 <> EntryID2 Then 
 Debug.Print "Error obtaining EntryID for " &; oMsg.Subject 
 End If 
 Next 
End Sub
```


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

