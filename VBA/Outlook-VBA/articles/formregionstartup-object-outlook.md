---
title: FormRegionStartup Object (Outlook)
keywords: vbaol11.chm3213
f1_keywords:
- vbaol11.chm3213
ms.prod: outlook
ms.assetid: 948ea6b7-2962-57e7-618d-fa0977b65651
ms.date: 06/08/2017
---


# FormRegionStartup Object (Outlook)

Defines an interface that allows an add-in to specify the storage and the user interface of a form region, obtains an object for that form region, and determines when the form region is about to be displayed in a form or in the Reading Pane.


## Remarks

The  **FormRegionStartup** interface is an abstract class, which means that it cannot be instantiated directly. In Visual Basic, you can use the **Implements** keyword to provide the methods of **FormRegionStartup** in your add-in class as follows:


```
Implements Outlook.FormRegionStartup
```

An add-in deploying a form region in an Outlook form on a client computer must implement the  **FormRegionStartup** interface which consists of the two methods, **[BeforeFormRegionShow](formregionstartup-beforeformregionshow-method-outlook.md)** and **[GetFormRegionStorage](formregionstartup-getformregionstorage-method-outlook.md)**. When Outlook loads the add-in, Outlook queries the **IDTExtensibility2** interface for **FormRegionStartup**.

The add-in indicates the storage and layout file for the form region in  **GetFormRegionStorage**. By calling **GetFormRegionStorage**, Outlook allocates storage and calculates the layout for the form region, instantiates an object for the form region, and returns a value representing the storage allocated to the add-in. If **GetFormRegionStorage** is successful, just before the form region is displayed in an Inspector window or in the Reading Pane, Outlook will call **BeforeFormRegionShow** and pass the **[FormRegion](formregion-object-outlook.md)** object of the form region to the add-in. The add-in uses this opportunity before the form region is displayed to update any controls in the form region.

When the add-in closes the frame for the form region, the add-in must release the object for the form region.

For more information on writing add-ins for form regions, see [Extending a Form Region with an Add-in](http://msdn.microsoft.com/library/b1a28a20-a0b8-cc57-7672-da51ec8bb097%28Office.15%29.aspx). For examples of add-ins in C# and Visual Basic .NET that implement  **FormRegionStartup**, see code sample downloads on MSDN.


## Methods



|**Name**|
|:-----|
|[BeforeFormRegionShow](formregionstartup-beforeformregionshow-method-outlook.md)|
|[GetFormRegionIcon](formregionstartup-getformregionicon-method-outlook.md)|
|[GetFormRegionManifest](formregionstartup-getformregionmanifest-method-outlook.md)|
|[GetFormRegionStorage](formregionstartup-getformregionstorage-method-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
