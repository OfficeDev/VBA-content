---
title: PropertyAccessor.GetProperty Method (Outlook)
keywords: vbaol11.chm1970
f1_keywords:
- vbaol11.chm1970
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.GetProperty
ms.assetid: a5f3493b-f302-c7b6-f442-23a7605be1c1
ms.date: 06/08/2017
---


# PropertyAccessor.GetProperty Method (Outlook)

Returns an  **Object** that represents the value of the property specified by _SchemaName_ .


## Syntax

 _expression_ . **GetProperty**( **_SchemaName_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaName_|Required| **String**|The name of the property whose value is to be returned. The property is referenced by namespace. For more information, see [Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|

### Return Value

A  **Variant** value that represents the value of the requested property as specified by _SchemaName_ .


## Remarks

The type of the return value will be the same as the type of the underlying property. Certain raw property types such as  **PT_OBJECT** are unsupported and will raise an error. If you require conversion of the raw property type, for example, from **PT_BINARY** to a string, or from **PT_SYSTIME** to a local time, use the helper methods[PropertyAccessor.BinaryToString](propertyaccessor-binarytostring-method-outlook.md) and[PropertyAccessor.UTCToLocalTime](propertyaccessor-utctolocaltime-method-outlook.md). 

For more information on getting properties using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## Example

The following code sample demonstrates how to use the  **GetProperty** method to read a MAPI property that belongs to a **[MailItem](mailitem-object-outlook.md)** but which is not exposed in the Outlook object model, **PR_TRANSPORT_MESSAGE_HEADERS** .


```vb
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


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

