---
title: PropertyAccessor.SetProperties Method (Outlook)
keywords: vbaol11.chm1973
f1_keywords:
- vbaol11.chm1973
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.SetProperties
ms.assetid: bf7c86da-5146-9567-5b7e-3e5e63ee5587
ms.date: 06/08/2017
---


# PropertyAccessor.SetProperties Method (Outlook)

Sets the properties specified by the array  _SchemaNames_ to the values specified by the array _Values_ .


## Syntax

 _expression_ . **SetProperties**( **_SchemaNames_** , **_Values_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaNames_|Required| **Variant**|An array of names of properties whose values are to be set as specified by the  _Values_ parameter. These properties are referenced by namespace. For more information, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|
| _Values_|Required| **Variant**|An array of values that are to be set for the properties specified by the  _SchemaNames_ parameter.|

### Return Value

A  **Variant** that is **Null** ( **Nothing** in VBA) if the operation is successful. If there is an error before any properties are set, for example, the number of elements in the _SchemaNames_ array does not match that in the _Values_ array, and an **Err** value will be returned. If there is an error during the setting of the properties, the return value is an array of **Err** objects, with the number of elements in this array being the same as that of the _SchemaNames_ array. An **Err** value in the array is mapped to the error result of setting the corresponding property in the _SchemaNames_ parameter.


## Remarks

If the property does not exist and the  _SchemaNames_ element contains a valid property specifier, then **SetProperties** creates the property and assigns the property with the value specified by _Values_ . The type of the property will be the type of the element passed in _Values_ . If the property does exist, then **SetProperties** assigns the property the value as specified by _Values_ .

Note that a custom property created by using the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** is not supported in a custom view. If you want to view a custom property on an item, create the property by using the **[Add](userproperties-add-method-outlook.md)** method of the **[UserProperties](userproperties-object-outlook.md)** object.

If the parent object of the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** supports an explicit **Save** operation, then the properties should be saved to the object with an explicit **Save** method call. If the object does not support an explicit **Save** operation, then the properties are saved to the object when **SetProperties** is called.

Use caution and ensure that all exceptions are handled correctly. Conditions where setting properties fails include:


- The property is read-only, as some Outlook and MAPI properties are read-only.
    
- The property referenced by the specified namespace is not found.
    
- The property is specified in an invalid format and cannot be parsed.
    
- The property does not exist and cannot be created.
    
- The property exists but is passed a value of an incorrect type.
    
- Cannot open the property because the client is offline.
    
- The property is created using the  **[UserProperties.Add](userproperties-add-method-outlook.md)** method. When setting the property for the first time, you must use the **[UserProperty.Value](userproperty-value-property-outlook.md)** property instead of the **SetProperties** or **[SetProperty](propertyaccessor-setproperty-method-outlook.md)** method of the **PropertyAccessor** object.
    


For more information on setting properties using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## Example

This code sample demonstrates how the  **SetProperties** method sets the values of multiple properties. If a property does not exist, then **SetProperties** will create the property as long as the parent object supports the creation of those properties. Since the **[MailItem](mailitem-object-outlook.md)** object supports a **[MailItem.Save](mailitem-save-method-outlook.md)** operation, the properties here are saved with an explicit `oMail.Save`.


```vb
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
 prop1 = "http://schemas.microsoft.com/mapi/string/" &; _ 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mylongprop" 
 prop2 = "http://schemas.microsoft.com/mapi/string/" &; _ 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mystringprop" 
 prop3 = "http://schemas.microsoft.com/mapi/string/" &; _ 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/mydateprop" 
 prop4 = "http://schemas.microsoft.com/mapi/string/" &; _ 
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


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

