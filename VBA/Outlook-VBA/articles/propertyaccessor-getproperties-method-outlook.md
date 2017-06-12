---
title: PropertyAccessor.GetProperties Method (Outlook)
keywords: vbaol11.chm1972
f1_keywords:
- vbaol11.chm1972
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.GetProperties
ms.assetid: f1ba3c52-428a-9e9f-5b81-b68c5f27aa0f
ms.date: 06/08/2017
---


# PropertyAccessor.GetProperties Method (Outlook)

Obtains the values of the properties specified by the one-dimensional array  _SchemaNames_ .


## Syntax

 _expression_ . **GetProperties**( **_SchemaNames_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaNames_|Required| **Variant**|An array that contains the names of the properties whose values are to be returned. These properties are referenced by namespace. For more information, see [Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|

### Return Value

A Variant that represents an array of values of the properties specified in the parameter  _SchemaNames_ . The number of elements in the returned array equals the number of elements in the _SchemaNames_ array. If an error occurs for getting a specific property, the **Err** value will be returned in the corresponding location in the returned array.


## Remarks

The array returned by  **GetProperties** can contain elements of different types, depending on the type of the property requested. The type of the array element returned by **GetProperties** will be the same as the type of the underlying property. Certain raw property types such as **PT_OBJECT** are unsupported and will raise an error. If you require conversion of the raw property type, for example, from **PT_BINARY** to a string, or from **PT_SYSTIME** to a local time, use the helper methods[PropertyAccessor.BinaryToString](propertyaccessor-binarytostring-method-outlook.md) and[PropertyAccessor.UTCToLocalTime](propertyaccessor-utctolocaltime-method-outlook.md). 

For more information on getting properties using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## Example

The following code sample shows how to use the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object to get MAPI properties that are not exposed on an Outlook item, namely: **PR_SUBJECT** , **PR_ATTR_HIDDEN** , **PR_ATTR_READONLY** , and **PR_ATTR_SYSTEM** . This code sample uses the **GetProperties** method to retrieve them in a single call, specifying an array of namespace references to these properties, and obtains a returned array that contains the raw value for each property.


```vb
Sub DemoPropertyAccessorGetProperties() 
 
 Dim PropNames() As Variant 
 
 Dim myValues As Variant 
 
 Dim i As Integer 
 
 Dim j As Integer 
 
 Dim oMail As Object 
 
 Dim oPA As Outlook.PropertyAccessor 
 
 
 
 'Get first item in the inbox 
 
 Set oMail = _ 
 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 
 'PR_SUBJECT, PR_ATTR_HIDDEN, PR_ATTR_READONLY, PR_ATTR_SYSTEM 
 
 PropNames = _ 
 
 Array("http://schemas.microsoft.com/mapi/proptag/0x0037001E", _ 
 
 "http://schemas.microsoft.com/mapi/proptag/0x10F4000B", _ 
 
 "http://schemas.microsoft.com/mapi/proptag/0x10F6000B", _ 
 
 "http://schemas.microsoft.com/mapi/proptag/0x10F5000B") 
 
 'Obtain an instance of a PropertyAccessor object 
 
 Set oPA = oMail.PropertyAccessor 
 
 'Get myValues array with GetProperties call 
 
 myValues = oPA.GetProperties(PropNames) 
 
 For i = LBound(myValues) To UBound(myValues) 
 
 'Examine the type of the element 
 
 If IsError(myValues(i)) Then 
 
 'CVErr returns a variant of subtype error 
 
 Debug.Print (CVErr(myValues(i))) 
 
 ElseIf IsArray(myValues(i)) Then 
 
 propArray = myValues(i) 
 
 For j = LBound(propArray) To UBound(propArray) 
 
 Debug.Print (propArray(j)) 
 
 Next 
 
 ElseIf IsNull(myValues(i)) Then 
 
 Debug.Print ("Null value") 
 
 ElseIf IsEmpty(myValues(i)) Then 
 
 Debug.Print ("Empty value") 
 
 ElseIf IsDate(myValues(i)) Then 
 
 Debug.Print (oPA.UTCToLocalTime(myValues(i))) 
 
 'VB does not have IsBinary function 
 
 ElseIf VarType(myValues(i)) = vbByte Then 
 
 Debug.Print (oPA.BinaryToString(myValues(i))) 
 
 Else 
 
 Debug.Print (myValues(i)) 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

