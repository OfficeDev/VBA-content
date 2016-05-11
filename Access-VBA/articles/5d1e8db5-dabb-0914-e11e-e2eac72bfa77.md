
# Microsoft OLE DB Provider for Internet Publishing

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

The Microsoft OLE DB Provider for Internet Publishing allows ADO to access resources served by Microsoft FrontPage or Microsoft Internet Information Server. Resources include web source files such as HTML files, or Windows 2000 web folders.


## Connection String Parameters

To connect to this provider, set the  _Provider_ argument of the[ConnectionString](c67a7daf-258f-d99d-6475-a4aa98d1e99d.md) property to:


```
 
MSDAIPP.DSO 

```

This value can also be set or read using the [Provider](1b795f51-93d7-431c-b1fe-0db95f69a56a.md) property.


## Typical Connection String

A typical connection string for this provider is:


```
 
"Provider=MSDAIPP.DSO;Data Source=ResourceURL ;User ID=userName ;Password=userPassword ;" 

```

-or-




```
 
"URL=ResourceURL ;User ID=userName ;Password=userPassword ;" 

```

The string consists of these keywords:



|**Keyword**|**Description**|
|:-----|:-----|
|**Provider**|Specifies the OLE DB Provider for Internet Publishing.|
|**Data Source** -or- **URL**|Specifies the URL of a file or directory published in a Web Folder.|
|**User ID**|Specifies the user name.|
|**Password**|Specifies the user password.|
If you set the  _ResourceURL_ value from the "URL=" in the connection string to an invalid value, by default the Internet Publishing Provider raises a dialog box to prompt for a valid value. This is undesirable behavior for a component in the middle tier of an application, because it suspends program execution until the dialog box is cleared and the client appears to freeze because it has not received a response from the component.


 **Note**  If MSDAIPP.DSO is explicitly specified as the value of the provider, either with the  _Provider_ connection string keyword or the **Provider** property, you cannot use "URL=" in the connection string. If you do, an error will occur. Instead, simply specify the URL as shown in the topic[Using ADO with the OLE DB Provider for Internet Publishing](864e5ece-0f50-5d88-4c40-f951b2a2eded.md).

