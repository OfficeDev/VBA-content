---
title: Filtering Items
ms.prod: outlook
ms.assetid: 4038e042-1b07-5d18-18b0-c2b58c9c42da
ms.date: 06/08/2017
---


# Filtering Items

This topic describes the general rules for specifying properties in filters that are supported by various objects in Outlook. For more information about specifying conditions on properties to complete a filter, see the topics in the [Filter Syntax](#filter-syntax) section. 

A filter is a condition or a set of conditions that you can apply to a set of items to obtain a subset of items that meets the specified conditions. Outlook supports filters by using the Microsoft Jet query language syntax or the DAV Searching and Locating (DASL) syntax. Note that the Jet query language syntax has the same syntax as that supported by Microsoft Jet Expression Service, hence the name Jet query language.

As an example, you can filter contact items in your Contacts folder to obtain a list of contacts residing in Canada. In this case, you will be filtering on the **[HomeAddressCountry](contactitem-homeaddresscountry-property-outlook.md)** property. The filter, expressed as a Jet filter, will be `"[HomeAddressCountry] = 'Canada'"`.

Outlook provides filtering through the following entry points:


|Entry point|Jet filter support|DASL filter support|
|:-----|:-----|:-----|
|**[Application.AdvancedSearch](application-advancedsearch-method-outlook.md)**|No|Yes|
|**[Folder.GetTable](folder-gettable-method-outlook.md)**|Yes|Yes|
|**[Items.Find](items-find-method-outlook.md)**|Yes|Yes. Note that if you use the query keywords **ci_phrasematch** or **ci_startswith** in the filter, you will get an error.|
|**[Items.Restrict](items-restrict-method-outlook.md)**|Yes|Yes|
|**[Search.GetTable](search-gettable-method-outlook.md)**|No|Yes|
|**[Table.FindRow](table-findrow-method-outlook.md)**|Yes|Yes. Note that if you use the query keywords **ci_phrasematch** or **ci_startswith** in the filter, you will get an error.|
|**[Table.Restrict](table-restrict-method-outlook.md)**|Yes|Yes|
|**[View.Filter](view-filter-property-outlook.md)**|No|Yes|


 **Note**  A filter must contain a query in either Jet or DASL syntax but not a mixture of both.


## Property specifiers

When specifying properties in a Jet filter or DASL filter using any of the above entry points, follow these guidelines.


||**Jet filter**|**DASL filter**|
|:-----|:-----|:-----|
|**Applicable properties**|Most explicit built-in and custom item-level properties; see corresponding method topic for unsupported properties.|Most built-in and custom item-level properties with and without explicit string names; see corresponding method topic for unsupported properties.|
|**Referencing properties**|<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>By their explicit string names.</p></li><li><p>Explicit built-in properties can only be referenced by their names in English and not any other localized language.</p></li><li><p>Custom properties can be referenced by their names in English or  a localized language.</p></li></ul>|By their namespaces.|
|**Format of reference**|<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Enclose square brackets ('['']') around explicit string names.</p></li><li><p>Property names are not case-sensitive.</p></li><li><p>Spaces are not allowed in explicit built-in properties.</p></li><li><p>Spaces are allowed in custom properties.</p></li></ul>|<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>All DASL queries begin with a case-sensitive prefix "@SQL=", with the exception of DASL queries for <b>Application.AdvancedSearch</b>.</p></li><li><p>Property referenced by namespace must be enclosed in double quotes.</p></li><li><p>Property referenced by namespace is case-sensitive.</p></li><li><p>If a space exists in the name of a custom property, the space must be replaced by "%20". In general, URL encoding applies the same way to characters in  a DASL query as in a URL.</p></li></ul>|
|**Error conditions**|Returns an error if a custom property in the filter is not defined, or the filter is empty, has an invalid argument, or cannot be parsed.|Returns an error if a custom property in the filter is not defined, or the filter is empty, has an invalid argument, or cannot be parsed.|


## Filter syntax

The syntax of a filter depends on the type of the property you are filtering on. The following topics provide further information about how to construct a filter based on a specific property type:

-  [Filtering a Custom Field](filtering-a-custom-field.md)
    
-  [Filtering Items Using a Boolean Comparison](filtering-items-using-a-boolean-comparison.md)
    
-  [Filtering Items Using a Comparison with a Keywords Property](filtering-items-using-a-comparison-with-a-keywords-property.md)
    
-  [Filtering Items Using a Date-time Comparison](filtering-items-using-a-date-time-comparison.md)
    
-  [Filtering Items Using a String Comparison](filtering-items-using-a-string-comparison.md)
    
-  [Filtering Items Using a Variable](filtering-items-using-a-variable.md)
    
-  [Filtering Items Using an Integer Comparison](filtering-items-using-an-integer-comparison.md)
    
-  [Filtering Items Using Comparison and Logical Operators](filtering-items-using-comparison-and-logical-operators.md)
    
-  [Filtering Items Using Query Keywords](filtering-items-using-query-keywords.md)
    



