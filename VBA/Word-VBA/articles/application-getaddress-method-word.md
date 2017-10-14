---
title: Application.GetAddress Method (Word)
keywords: vbawd10.chm158335298
f1_keywords:
- vbawd10.chm158335298
ms.prod: word
api_name:
- Word.Application.GetAddress
ms.assetid: b0081a05-be87-d0e4-31a6-b0aab02a3371
ms.date: 06/08/2017
---


# Application.GetAddress Method (Word)

Returns an address from the default address book.


## Syntax

_expression_. **GetAddress** (**_Name_**, **_AddressProperties_**, **_UseAutoText_**, **_DisplaySelectDialog_**, **_SelectDialog_**, **_CheckNamesDialog_**, **_RecentAddressesChoice_**, **_UpdateRecentAddresses_**)

_expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**Variant**|The name of the addressee, as specified in the **Search Name** dialog box in the address book.|
| _AddressProperties_|Optional|**Variant**|If _UseAutoText_ is **True**, this argument denotes the name of an AutoText entry that defines a sequence of address book properties. If _UseAutoText_ is **False** or omitted, this argument defines a custom layout.<br/><br/>Valid address book property names or sets of property names are surrounded by angle brackets (`"<" and ">"`) and separated by a space or a paragraph mark (for example, `"<PR_GIVEN_NAME> <PR_SURNAME>" &; vbCr &; "<PR_OFFICE_TELEPHONE_NUMBER>"`).<br/><br/>If the _AddressProperties_ parameter is omitted, a default AutoText entry named "AddressLayout" is used. If "AddressLayout" hasn't been defined, the following address layout definition is used: `"<PR_GIVEN_NAME> <PR_SURNAME>" &; vbCr &; "<PR_STREET_ADDRESS>" &; vbCr &; "<PR_LOCALITY>" &; ", " &; "<PR_STATE_OR_PROVINCE>" &; " " &; "<PR_POSTAL_CODE>" &; vbCr &; "<PR_COUNTRY>"`.<br/><br/>For a list of the valid address book property names, see the **AddAddress** method.|
| _UseAutoText_|Optional|**Variant**|**True** if _AddressProperties_ specifies the name of an AutoText entry that defines a sequence of address book properties; **False** if it specifies a custom layout.|
| _DisplaySelectDialog_|Optional|**Variant**|Specifies whether the **Select Name** dialog box is displayed, as shown in the [Results](#results) table.|
| _SelectDialog_|Optional|**Variant**|Specifies how the **Select Name** dialog box should be displayed (that is, in what mode), as shown in the [Display mode](#display-mode) table.|
| _CheckNamesDialog_|Optional|**Variant**|**True** to display the **Check Names** dialog box when the value of the _Name_ argument isn't specific enough.|
| _RecentAddressesChoice_|Optional|**Variant**|**True** to use the list of recently used return addresses.|
| _UpdateRecentAddresses_|Optional|**Variant**|**True** to add an address to the list of recently used addresses; **False** to not add the address. If _SelectDialog_ is set to 1 or 2, this argument is ignored.|

<br/>

#### Results

|**Value**|**Result**|
|:-----|:-----|
|0 (zero)|The **Select Name** dialog box isn't displayed.|
|1 or omitted|The **Select Name** dialog box is displayed.|
|2|The **Select Name** dialog box isn't displayed, and no search for a specific name is performed. The address returned by this method will be the previously selected address.|

<br/>

#### Display mode

|**Value**|**Display mode**|
|:-----|:-----|
|0 (zero) or omitted|Browse mode|
|1|Compose mode, with only the **To**: box|
|2|Compose mode, with both the **To**: and **CC**: boxes|

<br/>

### Return value

String

## Example

This example sets the variable _strAddress_ to John Smith's address, moves the insertion point to the beginning of the document, and inserts the address. The inserted address will include the default address book properties.

```vb
Dim strAddress 
 
strAddress = Application.GetAddress(Name:="John Smith", _ 
    CheckNamesDialog:=True) 
ActiveDocument.Range(Start:=0, End:=0).InsertAfter strAddress
```

The following example returns John Smith's address, using the "My Address Layout" AutoText entry as the layout definition. "My Address Layout" is defined in the active template and contains a set of address properties assigned to the text$ variable. The example also adds John Smith's address to the list of recently used addresses.

```vb
Dim TagIDArray(0 To 3) As String 
Dim ValueArray(0 To 3) As String 
Dim strAddress As String 
 
TagIDArray(0) = "PR_DISPLAY_NAME" 
TagIDArray(1) = "PR_GIVEN_NAME" 
TagIDArray(2) = "PR_SURNAME" 
TagIDArray(3) = "PR_COMMENT" 
ValueArray(0) = "Display_Name" 
ValueArray(1) = "John" 
ValueArray(2) = "Smith" 
ValueArray(3) = "This is a comment" 
 
Application.AddAddress TagID:=TagIDArray(), Value:=ValueArray() 
strAddress = Application.GetAddress(Name:="John Smith", _ 
    UpdateRecentAddresses:=True)
```


## See also

#### Concepts

- [Application Object](application-object-word.md)

