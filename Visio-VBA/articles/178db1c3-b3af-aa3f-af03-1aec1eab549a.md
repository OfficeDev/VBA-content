
# Application.RegisterRibbonX Method (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Registers the  ** [IRibbonExtensibility](http://msdn.microsoft.com/library/b27a7576-b6f5-031e-e307-78ef5f8507e0%28Office.15%29.aspx)** interface that is implemented by the specified add-on to populate the custom user interface (UI).


## Syntax

 _expression_. **RegisterRibbonX**( **_SourceAddOn_**,  **_TargetDocument_**,  **_TargetModes_**,  **_FriendlyName_**)

 _expression_A variable that represents an  ** [Application](5b3c8939-793f-116f-11b8-1d4170d95a63.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|SourceAddOn|Required| **IRibbonExtensibilty**|The add-on to register.|
|TargetDocument|Required| ** [Document](21640062-13a2-a2b2-7c61-7e707671207c.md)**|The document that uses the custom UI.|
|TargetModes|Required| ** [VisRibbonXModes](80f01121-3ea5-1ba8-bbea-ba81936ea4ae.md)**|The modes in which the custom UI should be visible. See Remarks for possible values.|
|FriendlyName|Required| **String**|The name to associate with the UI items and errors that originate in the add-on.|

### Return Value

 **Nothing**


## Remarks

If  _TargetDocument_ is null, the custom UI is defined at the level of the application. Otherwise, Microsoft Visio binds the visibility of the custom UI to the specified document. The UI does not appear in conjunction with any other document.

 _TargetModes_ can be one or more of the following **VisRibbonXModes** constants.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visRxModeNone**|0|Display the custom UI when no documents are open.|
| **visRxModeDrawing**|1|Display the custom UI in Drawing mode.|
| **visRxModeStencil**|2|Display the custom UI in Stencil mode.|
| **visRxModePrintPreview**|4|Display the custom UI in Print Preview mode.|
If  _FriendlyName_ is null, the method fails.

