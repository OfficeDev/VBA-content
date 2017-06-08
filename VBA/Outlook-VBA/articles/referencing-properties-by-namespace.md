---
title: Referencing Properties by Namespace
ms.prod: outlook
ms.assetid: c1c7bfa9-64d7-81d2-84e7-f0a4c57780b3
ms.date: 06/08/2017
---


# Referencing Properties by Namespace

This topic lists the namespaces that are supported by the  **PropertyAccessor**,  **Table**,  **View**, and their children objects,and discusses referencing named properties.


## Namespaces Used by Outlook Objects

The following table summarizes the namespaces and the Outlook objects that the namespaces support. Note that property references by namespaces are case-sensitive.



| **Namespaces**| **Supported Outlook Objects**|
|:-----|:-----|
|http://schemas.microsoft.com/mapi/proptag| [Outlook item objects](outlook-item-objects.md),  **[AddressEntry](addressentry-object-outlook.md)**,  **[AddressList](addresslist-object-outlook.md)**,  **[Attachment](attachment-object-outlook.md)**,  **[ExchangeDistributionList](exchangedistributionlist-object-outlook.md)**,  **[ExchangeUser](exchangeuser-object-outlook.md)**,  **[Folder](folder-object-outlook.md)**,  **[Recipient](recipient-object-outlook.md)**, and  **[Store](store-object-outlook.md)** objects.|
|http://schemas.microsoft.com/mapi/id| (Same as above)|
|http://schemas.microsoft.com/mapi/string|(Same as above)|
|http://schemas.microsoft.com/exchange|(Same as above)|
|urn:schemas-microsoft-com:office:office|Outlook item objects|
|urn:schemas-microsoft-com:office:outlook|Outlook item objects|
|DAV:|Outlook item objects|
|urn:schemas:calendar|Outlook item objects|
|urn:schemas:contacts|Outlook item objects|
|urn:schemas:httpmail|Outlook item objects|
|urn:schemas:mailheader|Outlook item objects|



## Messaging Application Programming Interface (MAPI) Namespaces

Many properties that Outlook supports are MAPI properties. The  **[PropertyAccessor](propertyaccessor-object-outlook.md)** object supports three subnamespaces of the MAPI namespace: proptag, id, and string. Each of the following sections contains a description for the subnamespace, a description for the format to reference a property in that subnamespace, and a definition of the syntax as expressed in Augmented Backus-Naur Form (ABNF), that is specified in [[RFC4234]](http://ietfreport.isoc.org/idref/rfc4234/).


-  **proptag** namespace
    
    This namespace is used to access properties in the MAPI namespace using the property tag of a property. It supports only properties in the MAPI property range (that is, properties with a property identifier below 0x8000). The following is the format to reference a property in this namespace:
    
     **http://schemas.microsoft.com/mapi/proptag/0xHHHHHHHH**
    
     **HHHHHHHH** represents a hexadecimal property tag value, with a unique property identifier in the higher-order 16 bits, and a property type in the lower-order 16 bits. Every MAPI property must have a property tag, regardless of whether the property is defined by MAPI, Outlook, or a service provider. The hexadecimal value must follow the prefix "0x". Formally, references of properties in this namespace can be defined in ABNF as follows:
    


```
  proptag-specifier = "http://schemas.microsoft.com/mapi/proptag/x" property-id property-type 
property-id = 4HEXDIG 
property-type = 4HEXDIG
```


    For example, the following represents the MAPI property  **PidTagSubject** that Outlook exposes in its object model as **Subject**: 
    
     **http://schemas.microsoft.com/mapi/proptag/0x0037001E**
    
-  **id** namespace
    
    This namespace is used to access properties in a namespace identified by the globally unique identifier (GUID) of the namespace, using the identifier of the property. The following is the format to reference a property in this namespace:
    
     **http://schemas.microsoft.com/mapi/id/{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}/HHHHHHHH**
    
     **{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}** represents the namespace GUID, and **HHHHHHHH** represents the property tag.
    
    Formally, references of properties in this namespace can be defined in ABNF as follows:
    


```
  id-specifier = "http://schemas.microsoft.com/mapi/id/" property-set "/x" property-long-id 
property-set = "{" 8HEXDIG "-" 4HEXDIG "-" 4HEXDIG "-" 4HEXDIG "-" 12HEXDIG "}" 
property-long-id = 8HEXDIG
```


    For example, the following represents the Outlook  **NoAging** property:
    
     **http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/850E000B**
    
-  **string** namespace
    
    This namespace is used to access string-named properties in an identified namespace. The following is the format to reference a property in this namespace:
    
     **http://schemas.microsoft.com/mapi/string/{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}/ _name_**
    
     **{HHHHHHHH-HHHH-HHHH-HHHH-HHHHHHHHHHHH}** represents the namespace GUID, and **_name_** is the local property name defined as a string.
    
    Formally, references of properties in this namespace can be defined in ABNF as follows:
    


```
  string-specifier = "http://schemas.microsoft.com/mapi/string/" property-set "/" property-name 
property-set = "{" 8*HEXDIG "-" 4*HEXDIG "-" 4*HEXDIG "-" 4*HEXDIG "-" 12*HEXDIG "}" 
property-name = 1*CHAR
```


    The following is an example that uses this namespace:
    
     **http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/content-class**
    
    Escaping rules apply to referencing named properties in the  **string** namespace. When referencing a named property that has a string identifier (for example, Author, Company, and Title), if the property name contains a space, single quote, double quote, or percent character, you must use Universal Resource Locator (URL) escaping and represent such characters with the corresponding escape string as shown in the following table.
    

| **Character in Property Reference**| **Escape String**|
|Space character|%20|
|Double quote|%22|
|Single quote|%27|
|Percent character|%25|

    The following is an example of how you specify and get the value of a named property,  **Mom's "Gift"**, defined in the MAPI string namespace, by using the  ** [PropertyAccessor.GetProperty](propertyaccessor-getproperty-method-outlook.md)** method:
    


```vb
  PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Mom%27s%20%22Gift%22")
```


## Exchange Namespace

The exchange namespace is used to access string-named Exchange properties. The following is the format to reference a property in this namespace:

 **http://schemas.microsoft.com/exchange/ _name_**

 **_name_** is the local property name defined as a string.

The following is an example of a property referenced by this namespace:

 **http://schemas.microsoft.com/exchange/readreceiptrequested**


## Office Namespaces

The  **PropertyAccessor** object supports two Office subnamespaces:


- Office namespace 
    
    This namespace is used to access properties of the  **[DocumentItem](documentitem-object-outlook.md)** object. The following is the format to reference a property in this namespace:
    
     **urn:schemas-microsoft-com:office:office# _name_**
    
     **_name_** is the local property name defined as a string.
    
    The following are some examples of referencing  **DocumentItem** properties using the Office namespace:
    
      -  **urn:schemas-microsoft-com:office:office#Subject**
    
  -  **urn:schemas-microsoft-com:office:office#Template**
    
- Outlook namespace
    
    This namespace is used to access Outlook item-level properties. Similar to other namespaces that support property referencing, use this namespace to access Outlook properties that are not explicitly exposed in the object model. The following is the format to reference a property in this namespace: 
    
     **urn:schemas-microsoft-com:office:outlook# _name_**
    
     **_name_** is the local property name defined as a string.
    
    The following is an example of referencing an Outlook item-level property by using the Outlook namespace: 
    
     **urn:schemas-microsoft-com:office:outlook#remotemessagesize**
    

## Distributed Authoring and Versioning (DAV) Namespaces

DAV namespaces are used to access Outlook item-level properties. A property in a DAV namespace is scoped using a Uniform Resource Identifier (URI) namespace reference. The format is a concatenation of the namespace URI prefix and the local property name expressed in a string, with the namespace URI being either a Uniform Resource Name (URN) or Uniform Resource Locator (URL).

The following are the DAV namespaces that the  **PropertyAccessor** object supports:


-  **DAV:**
    
-  **urn:schemas:calendar**
    
-  **urn:schemas:contacts**
    
-  **urn:schemas:httpmail**
    
-  **urn:schemas:mailheader**
    
These are some examples of properties being referenced by different DAV namespaces:


-  **DAV:checkintime**
    
-  **urn:schemas:httpmail:subject**
    
-  **urn:schemas:mailheader:subject**
    

## See also


#### Other resources


 [MAPI Property Tags](http://msdn.microsoft.com/library/380dad4c-7fbf-4c49-b67c-ab612c923499.aspx)<br>
 [MAPI Property Identifier Overview](http://msdn.microsoft.com/library/957aa00f-23d8-4f3b-bbc2-7d54f17b47b5.aspx)<br>
 [MAPI Property Type Overview](http://msdn.microsoft.com/library/b762f5fb-7c2c-4303-96f7-0b6e657146c9.aspx)<br>
 [Property Identifier Ranges](http://msdn.microsoft.com/library/c01e95bb-be25-490d-880b-60674f890258.aspx)<br>
 [Property Types](http://msdn.microsoft.com/library/71967150-1005-4c85-90f1-76fc7876c0d0.aspx)<br>
 [MAPI Named Properties](http://msdn.microsoft.com/library/464b1297-9d90-47bd-afc4-3dc63b106cb7.aspx)

