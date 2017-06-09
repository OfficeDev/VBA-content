---
title: Properties Object (Access)
keywords: vbaac10.chm10046
f1_keywords:
- vbaac10.chm10046
ms.prod: access
api_name:
- Access.Properties
ms.assetid: 7e888aad-e783-dfc5-46df-9d92c89cfc35
ms.date: 06/08/2017
---


# Properties Object (Access)

The  **Properties** collection contains all of the built-in properties in an instance of an open form, report, or control. These properties uniquely characterize that instance of the object.


## Remarks

Use the  **Properties** collection in[Visual Basic](http://msdn.microsoft.com/library/03349d86-f107-9e49-89df-62f55f3a0735%28Office.15%29.aspx)or in an expression to refer to form, report, or control properties on forms or reports that are currently open.

You can use the  **Properties** collection of an object to enumerate the object's built-in properties. You don't need to know beforehand exactly which properties exist or what their characteristics ( **Name** and **Value** properties) are to manipulate them.


 **Note**  In addition to the built-in properties, you can also create and add your own user-defined properties. To add a user-defined property to an existing instance of an object, see the  **[AccessObjectProperties](accessobjectproperties-object-access.md)** collection and **[Add](accessobjectproperties-add-method-access.md)** method topics.


## Example

The following example enumerates the  **Forms** collection and prints the name of each open form in the **Forms** collection. It then enumerates the **Properties** collection of each form and prints the name of each property and value.


```
Sub AllOpenForms() 
 Dim frm As Form, prp As Property 
 
 ' Enumerate Forms collection. 
 For Each frm In Forms 
 ' Print name of form. 
 Debug.Print frm.Name 
 ' Enumerate Properties collection of each form. 
 For Each prp In frm.Properties 
 ' Print name of each property. 
 Debug.Print prp.Name; " = "; prp.Value 
 Next prp 
 Next frm 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](properties-application-property-access.md)|
|[Count](properties-count-property-access.md)|
|[Item](properties-item-property-access.md)|
|[Parent](properties-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
