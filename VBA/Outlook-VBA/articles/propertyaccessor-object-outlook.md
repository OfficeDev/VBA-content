---
title: PropertyAccessor Object (Outlook)
keywords: vbaol11.chm3157
f1_keywords:
- vbaol11.chm3157
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor
ms.assetid: 2fc91e13-703c-3ec9-9066-ffee7144306c
ms.date: 06/08/2017
---


# PropertyAccessor Object (Outlook)

Provides the ability to create, get, set, and delete properties on objects.


## Remarks

Use the  **PropertyAccessor** object to get and set item-level properties that are not explicitly exposed in the Outlook object model, or properties for the following non-item objects: **[AddressEntry](addressentry-object-outlook.md)**, **[AddressList](addresslist-object-outlook.md)**, **[Attachment](http://msdn.microsoft.com/library/3e11582b-ac90-0948-bc37-506570bb287b%28Office.15%29.aspx)**, **[ExchangeDistributionList](http://msdn.microsoft.com/library/2830dfba-6c0a-a81f-6b98-92ac2aafb59d%28Office.15%29.aspx)**, **[ExchangeUser](exchangeuser-object-outlook.md)**, **[Folder](folder-object-outlook.md)**, **[Recipient](recipient-object-outlook.md)**, and **[Store](store-object-outlook.md)**.

To get or set multiple custom properties, use the  **PropertyAccessor** object instead of the **[UserProperties](userproperties-object-outlook.md)** object for better performance.

For more information on using the  **PropertyAccessor** object, see[Properties Overview](http://msdn.microsoft.com/library/242c9e89-a0c5-ff89-0d2a-410bd42a3461%28Office.15%29.aspx).


## Example

The following code sample demonstrates how to use the  **[PropertyAccessor.GetProperty](http://msdn.microsoft.com/library/a5f3493b-f302-c7b6-f442-23a7605be1c1%28Office.15%29.aspx)** method to read a MAPI property that belongs to a **[MailItem](http://msdn.microsoft.com/library/14197346-05d2-0250-fa4c-4a6b07daf25f%28Office.15%29.aspx)** but that is not exposed in the Outlook object model, **PR_TRANSPORT_MESSAGE_HEADERS**.


```
Sub DemoPropertyAccessorGetProperty() 
 
 Dim PropName, Header As String 
 
 Dim oMail As Object 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'PR_TRANSPORT_MESSAGE_HEADERS 
 
 PropName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E" 
 
 'Obtain an instance of PropertyAccessor class 
 
 Set oPA = oMail.PropertyAccessor 
 
 'Call GetProperty 
 
 Header = oPA.GetProperty(PropName) 
 
 Debug.Print (Header) 
 
End Sub
```

The next code sample demonstrates how the  **[PropertyAccessor.SetProperties](http://msdn.microsoft.com/library/bf7c86da-5146-9567-5b7e-3e5e63ee5587%28Office.15%29.aspx)** method sets the values of multiple properties. If a property does not exist, then **SetProperties** will create the property as long as the parent object supports the creation of those properties. If the object supports an explicit **Save** operation, then the properties are saved to the object when the explicit **Save** operation is called. If the object does not support an explicit **Save** operation, then the properties are saved to the object when **SetProperties** is called.




```
Sub DemoPropertyAccessorSetProperties() 
 
 Dim PropNames(), myValues() As Variant 
 
 Dim arrErrors As Variant 
 
 Dim prop1, prop2, prop3, prop4 As String 
 
 Dim i As Integer 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'Names for properties using the MAPI string namespace 
 
 prop1 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop" 
 
 prop2 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop" 
 
 prop3 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop" 
 
 prop4 = "http://schemas.microsoft.com/mapi/string/" &amp; _ 
 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myboolprop" 
 
 PropNames = Array(prop1, prop2, prop3, prop4) 
 
 myValues = Array(1020, "111-222-Kudo", Now(), False) 
 
 'Set values with SetProperties call 
 
 'If the properties do not exist, then SetProperties 
 
 'adds the properties to the object when saved. 
 
 'The type of the property is the type of the element 
 
 'passed in myValues array. 
 
 Set oPA = oMail.PropertyAccessor 
 
 arrErrors = oPA.SetProperties(PropNames, myValues) 
 
 If Not (IsEmpty(arrErrors)) Then 
 
 'Examine the arrErrors array to determine if any 
 
 'elements contain errors 
 
 For i = LBound(arrErrors) To UBound(arrErrors) 
 
 'Examine the type of the element 
 
 If IsError(arrErrors(i)) Then 
 
 Debug.Print (CVErr(arrErrors(i))) 
 
 End If 
 
 Next 
 
 End If 
 
 'Save the item 
 
 oMail.Save 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[BinaryToString](http://msdn.microsoft.com/library/4a3801af-0a7c-4b8a-7367-600c09047b28%28Office.15%29.aspx)|
|[DeleteProperties](http://msdn.microsoft.com/library/e9c11799-cb75-fd8c-0c98-aca46796bb46%28Office.15%29.aspx)|
|[DeleteProperty](http://msdn.microsoft.com/library/9acb52b5-13a7-7363-7e17-83804037f33b%28Office.15%29.aspx)|
|[GetProperties](http://msdn.microsoft.com/library/f1ba3c52-428a-9e9f-5b81-b68c5f27aa0f%28Office.15%29.aspx)|
|[GetProperty](http://msdn.microsoft.com/library/a5f3493b-f302-c7b6-f442-23a7605be1c1%28Office.15%29.aspx)|
|[LocalTimeToUTC](http://msdn.microsoft.com/library/c19f60b2-441f-77b3-eb83-9cfd899e3a52%28Office.15%29.aspx)|
|[SetProperties](http://msdn.microsoft.com/library/bf7c86da-5146-9567-5b7e-3e5e63ee5587%28Office.15%29.aspx)|
|[SetProperty](http://msdn.microsoft.com/library/2a97c11d-3f5f-65fe-23d6-8efa40dca303%28Office.15%29.aspx)|
|[StringToBinary](http://msdn.microsoft.com/library/1ea95601-a21f-47d2-7a3c-166c4984fc25%28Office.15%29.aspx)|
|[UTCToLocalTime](http://msdn.microsoft.com/library/a56311ac-60ac-4f51-5255-d6840bf6004d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/51df74aa-6120-519b-3b68-e86e11222264%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/ef4c4ec9-8e80-34de-7699-be1defe52d7c%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/764b07a0-2bfa-1457-b587-bc2559ff72a1%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/db33aa4e-ad96-2db8-de9d-7aa9dd1a137f%28Office.15%29.aspx)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
[PropertyAccessor Object Members](http://msdn.microsoft.com/library/3356e345-8878-0ed7-6783-1e49ddecc066%28Office.15%29.aspx)
