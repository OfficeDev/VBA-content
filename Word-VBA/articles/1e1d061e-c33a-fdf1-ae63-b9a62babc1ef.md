
# Document.SendFaxOverInternet Method (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sends a document to a fax service provider, who faxes the document to one or more specfied recipients.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **SendFaxOverInternet**( **_Recipients_**,  **_Subject_**,  **_ShowMessage_**)

 _expression_Required. A variable that represents a  ** [Document](8d83487a-2345-a036-a916-971c9db5b7fb.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Recipients|Optional| **Variant**|A  **String** that represents the fax numbers and e-mail addresses of the people to whom to send the fax. Separate multiple recipients with a semicolon.|
|Subject|Optional| **Variant**|A  **String** that represents the subject line for the faxed document.|
|ShowMessage|Optional| **Variant**| **True** displays the fax message before sending it. **False** sends the fax without displaying the fax message.|

## Remarks
<a name="sectionSection1"> </a>

Using the  **SendFaxOverInternet** method requires that a fax service is enabled on a user's computer. If a fax service is not enabled, the **SendFaxOverInternet** method will cause a runtime error.

The format used for specifying fax numbers in the Recipients parameter is either recipientsfaxnumber@usersfaxprovider or recipientsname@recipientsfaxnumber. You can access the user's fax provider information using the following registry path:




```
HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax
```

Use the FaxAddress key value at this registry location to determine the format to use for a user. If this registry entry does not exist, no fax service is available.


## Example
<a name="sectionSection2"> </a>

The following example sends a fax to the fax service provider, who will fax the message to the recipient.


```
ActiveDocument.SendFaxOverInternet _ 
 "14255550101@consolidatedmessenger.com", _ 
 "For your review", True
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Document Object](8d83487a-2345-a036-a916-971c9db5b7fb.md)
#### Other resources


 [Document Object Members](fc9ab457-0888-f917-3d52-387168ac23b9.md)
