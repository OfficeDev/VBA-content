---
title: FormRegionStartup.GetFormRegionStorage Method (Outlook)
keywords: vbaol11.chm2946
f1_keywords:
- vbaol11.chm2946
ms.prod: outlook
api_name:
- Outlook.FormRegionStartup.GetFormRegionStorage
ms.assetid: 685b5ed7-dd19-9040-664f-5deff6e738c7
ms.date: 06/08/2017
---


# FormRegionStartup.GetFormRegionStorage Method (Outlook)

Obtains appropriate storage for a form region based on the specified information.


## Syntax

 _expression_ . **GetFormRegionStorage**( **_FormRegionName_** , **_Item_** , **_LCID_** , **_FormRegionMode_** , **_FormRegionSize_** )

 _expression_ A variable that represents an object that implements the **FormRegionStartup** interface.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FormRegionName_|Required| **String**|The internal name of the form region. This can be indicated by the <name> tag in the corresponding form region XML manifest.|
| _Item_|Required| **Object**|The Outlook item object that caused the loading of the form region.|
| _LCID_|Required| **Long**|The current locale ID.|
| _FormRegionMode_|Required| **[OlFormRegionMode](olformregionmode-enumeration-outlook.md)**|The mode that the form region is being loaded into.|
| _FormRegionSize_|Required| **[OlFormRegionSize](olformregionsize-enumeration-outlook.md)**|The type of form region being loaded, either adjoining or separate.|

### Return Value

A  **Variant** object representing the storage that Outlook has allocated for the form region. The type of the return value can be: **String** representing that the return value is a local path to an Outlook Form Storage (.OFS) file; **Byte()** representing that the return value is an array of bytes that contains the contents of the .OFS file; **IStorage** representing that the return value is a COM storage object **IStorage** (for C++ only); **Nothing** or **Null** , representing that Outlook could not allocate storage for this form region and will not load the form region.


## Remarks

The add-in must check for the return value of  **GetFormRegionStorage** . A form region will not load if any of the following is true of the returned storage:


- The returned storage is a .OFS file specified with a non-local path.
    
- The returned storage is a file but is not an .OFS file saved from the forms designer.
    


For examples of add-ins in C# and Visual Basic .NET that implement the  **[FormRegionStartup](formregionstartup-object-outlook.md)** interface, see code sample downloads on MSDN.


## See also


#### Concepts


[FormRegionStartup Interface](formregionstartup-object-outlook.md)

