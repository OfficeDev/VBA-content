---
title: PropertyAccessor.SetProperty Method (Outlook)
keywords: vbaol11.chm1971
f1_keywords:
- vbaol11.chm1971
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.SetProperty
ms.assetid: 2a97c11d-3f5f-65fe-23d6-8efa40dca303
ms.date: 06/08/2017
---


# PropertyAccessor.SetProperty Method (Outlook)

Sets the property specified by  _SchemaName_ to the value specified by _Value_ .


## Syntax

 _expression_ . **SetProperty**( **_SchemaName_** , **_Value_** )

 _expression_ A variable that represents a **PropertyAccessor** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SchemaName_|Required| **String**|The name of a property whose value is to be set as specified by the  _Value_ parameter. The property is referenced by namespace. For more information, see[Referencing Properties by Namespace](http://msdn.microsoft.com/library/c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3%28Office.15%29.aspx).|
| _Value_|Required| **Variant**|The value that is to be set for the property specified by the  _SchemaName_ parameter.|

## Remarks

If the property does not exist and the  _SchemaName_ contains a valid property specifier, then **SetProperty** creates the property and assigns the value specified by _Value_ . If the property does exist and _SchemaName_ is valid, then **SetProperty** assigns the property with the value specified by _Value_ .

Note that a custom property created by using the  **[PropertyAccessor](propertyaccessor-object-outlook.md)** is not supported in a custom view. If you want to view a custom property on an item, create the property by using the **[Add](userproperties-add-method-outlook.md)** method of the **[UserProperties](userproperties-object-outlook.md)** object.

If the parent object of the  **PropertyAccessor** supports an explicit **Save** operation, then the properties should be saved to the object with an explicit **Save** method call. If the object does not support an explicit **Save** operation, then the properties are saved to the object when **SetProperties** is called.

Use caution and ensure that all exceptions are handled correctly. Conditions where setting properties fails include:


- The property is read-only, as some Outlook and MAPI properties are read-only.
    
- The property referenced by the specified namespace is not found.
    
- The property is specified in an invalid format and cannot be parsed.
    
- The property does not exist and cannot be created.
    
- The property exists but is passed a value of an incorrect type.
    
- Cannot open the property because the client is offline.
    
- The property is created using the  **UserProperties.Add** method. When setting the property for the first time, you must use the **[UserProperty.Value](userproperty-value-property-outlook.md)** property instead of the **[SetProperties](propertyaccessor-setproperties-method-outlook.md)** or **SetProperty** method of the **PropertyAccessor** object.
    


For more information on setting properties using the  **PropertyAccessor** object, see[Best Practices for Getting and Setting Properties](http://msdn.microsoft.com/library/ec087bf8-cfac-9b20-3cb2-3bd308c5c63d%28Office.15%29.aspx).


## Example

The following code sample shows how to use the  **PropertyAccessor** to set a custom property on a **MailItem** object to a value. If the custom property does not exist, **PropertyAccessor.SetProperty** will create and then set the property. The property is saved with the **[MailItem.Save](mailitem-save-method-outlook.md)** method.


```vb
Sub DemoPropertyAccessorSetProperty() 
 Dim myProp As String 
 Dim myValue As Variant 
 Dim oMail As Outlook.MailItem 
 Dim oPA As Outlook.PropertyAccessor 
 'Get first item in the inbox 
 Set oMail = _ 
 Application.Session.GetDefaultFolder(olFolderInbox).Items(1) 
 'Name for custom property using the MAPI string namespace 
 myProp = "http://schemas.microsoft.com/mapi/string/" &; _ 
 "{FFF40745-D92F-4C11-9E14-92701F001EB3}/myCustomer" 
 myValue = "Dan Wilson" 
 'Set value with SetProperty call 
 'If the property does not exist, then SetProperty 
 'adds the property to the object when saved. 
 'The type of the property is the type of the element 
 'passed in myValue. 
 On Error GoTo ErrTrap 
 Set oPA = oMail.PropertyAccessor 
 oPA.SetProperty myProp, myValue 
 
 'Save the item 
 oMail.Save 
 Exit Sub 
ErrTrap: 
 Debug.Print Err.Number, Err.Description 
End Sub
```


## See also


#### Concepts


[PropertyAccessor Object](propertyaccessor-object-outlook.md)

