---
title: Filtering a Custom Field
ms.prod: outlook
ms.assetid: 36c0e15a-775d-5ce3-8e61-2a6bd305a746
ms.date: 06/08/2017
---


# Filtering a Custom Field

You can specify custom properties in filters using Microsoft Jet syntax or DAV Searching and Locating (DASL) syntax. The custom properties must be defined in the folder where you are applying the filter. If the custom properties are only defined in the item, the search will fail.


## Jet Queries

Custom properties can contain spaces in the property name. In a Jet query, as in all property name references, simply enclose the custom property name in square brackets. For example, the following Jet query retrieves all contacts where the custom property named "Preferred Gift" is exactly "Diamonds". For the query to succeed, the custom property named "Preferred Gift" has been defined in the folder that contains the custom contact items: 


```
criteria = "[Preferred Gift] = 'Diamonds'"
```


## DASL Queries

In a DASL query, if the name of a custom property contains spaces, you must apply Uniform Resource Locator (URL) encoding to each space character and replace the space with "%20". In general, URL encoding applies the same way to characters in a DASL query as in a URL.

When you construct a DASL query for a custom property, you must use the namespace GUID for Outlook custom properties in the following format: 

 **http://schemas.microsoft.com/mapi/string/{GUID}/PropertyName**

where  **{GUID}** is the following GUID:

 **{00020329-0000-0000-C000-000000000046}**


## Filtering Custom Properties Referenced by the MAPI String Namespace

If the custom property you are filtering for does not exist in the  **[UserDefinedProperties](userdefinedproperties-object-outlook.md)** collection for the folder, and if you are referencing the custom property by the MAPI string namespace, then you must explicitly append a type specifier to the namespace representation of the custom property. Note that you need to specify the type only when applying a DASL filter to search and filter entry points in the **[Items](items-object-outlook.md)** collection and the **[Table](table-object-outlook.md)** object, and to the **[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)** method.


 **Note**  The hexagonal type specifier must be of the form 0000HHHH with only 8 digits as opposed to 9. For more information on the hexagonal type specifiers (HHHH) for various MAPI types, see  [Property Types](http://msdn.microsoft.com/library/71967150-1005-4c85-90f1-76fc7876c0d0.aspx).

For example, if you want to use  **[Items.Restrict](items-restrict-method-outlook.md)** to search for the custom Unicode string property named "MyProperty" and this property does not exist in the **UserDefinedProperties** collection for the folder, you must append the Unicode string type specifier, 0000001f, to the representation of the property in the MAPI string namespace:




```
criteria = "@SQL=" &; Chr$(34) &; "http://schemas.microsoft.com/mapi/string/" _ 
&; "{00020329-0000-0000-C000-000000000046}/MyProperty"_ 
&; "/0000001f" &; Chr(34) &; " = '12-74440'" 

```


