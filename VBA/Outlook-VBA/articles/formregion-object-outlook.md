---
title: FormRegion Object (Outlook)
keywords: vbaol11.chm3018
f1_keywords:
- vbaol11.chm3018
ms.prod: outlook
api_name:
- Outlook.FormRegion
ms.assetid: 3a0b83eb-4076-9cb3-86a9-68f9e44df89f
ms.date: 06/08/2017
---


# FormRegion Object (Outlook)

Represents a form region in an Outlook form.


## Remarks

The  **FormRegion** object allows an add-in to add code behind a form region in a custom form to modify the appearance and behavior of the form region.

To obtain an instance of the  **FormRegion** object, an add-in must implement the **[FormRegionStartup](formregionstartup-object-outlook.md)** interface. Outlook allocates storage for the form region, instantiates an instance of the **FormRegion** object, and returns the **FormRegion** object in the **[GetFormRegionStorage](formregionstartup-getformregionstorage-method-outlook.md)** method.

When the add-in closes the frame for the form region, the add-in must release the object for the form region.

For more infomation on programming a form region, see [Extending a Form Region with an Add-in](http://msdn.microsoft.com/library/b1a28a20-a0b8-cc57-7672-da51ec8bb097%28Office.15%29.aspx).


## Events



|**Name**|
|:-----|
|[Close](formregion-close-event-outlook.md)|
|[Expanded](formregion-expanded-event-outlook.md)|

## Methods



|**Name**|
|:-----|
|[Reflow](formregion-reflow-method-outlook.md)|
|[Select](formregion-select-method-outlook.md)|
|[SetControlItemProperty](formregion-setcontrolitemproperty-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Application](formregion-application-property-outlook.md)|
|[Class](formregion-class-property-outlook.md)|
|[Detail](formregion-detail-property-outlook.md)|
|[DisplayName](formregion-displayname-property-outlook.md)|
|[EnableAutoLayout](formregion-enableautolayout-property-outlook.md)|
|[Form](formregion-form-property-outlook.md)|
|[FormRegionMode](formregion-formregionmode-property-outlook.md)|
|[Inspector](formregion-inspector-property-outlook.md)|
|[InternalName](formregion-internalname-property-outlook.md)|
|[IsExpanded](formregion-isexpanded-property-outlook.md)|
|[Item](formregion-item-property-outlook.md)|
|[Language](formregion-language-property-outlook.md)|
|[Parent](formregion-parent-property-outlook.md)|
|[Session](formregion-session-property-outlook.md)|
|[SuppressControlReplacement](formregion-suppresscontrolreplacement-property-outlook.md)|
|[Visible](formregion-visible-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
