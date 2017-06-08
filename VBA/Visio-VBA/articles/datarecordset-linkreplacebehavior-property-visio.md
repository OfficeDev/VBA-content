---
title: DataRecordset.LinkReplaceBehavior Property (Visio)
keywords: vis_sdr.chm16460365
f1_keywords:
- vis_sdr.chm16460365
ms.prod: visio
api_name:
- Visio.DataRecordset.LinkReplaceBehavior
ms.assetid: a49a9a44-1067-dfc6-0fb0-aee15064078b
ms.date: 06/08/2017
---


# DataRecordset.LinkReplaceBehavior Property (Visio)

Gets or sets how existing links between shapes and data rows are handled when methods that link shapes to data is called. Read/write.


 **Note**  This Visio object or member is available only to licensed users of Visio Professional 2013.


## Syntax

 _expression_ . **LinkReplaceBehavior**

 _expression_ An expression that returns a **DataRecordset** object.


### Return Value

VisLinkReplaceBehavior


## Remarks

The following constants for link replacement behaviors are declared by the Visio type library in  **VisLinkReplaceBehavior** :



|**Constant**|**Value **|**Description**|
|:-----|:-----|:-----|
| **visLinkReplaceAlways**|1|Always replace links when linking to a shape that has existing links|
| **visLinkReplaceNever**|0|Never replace links when linking to a shape that has existing links|
| **visLinkReplacePrompt**|2|Prompt the user before replacing links when the user attempts to create links in the Visio user interface (UI). |
These options correspond to those available in the  **Properties** dialog box for the tab corresponding to the data recordset in the **External Data** window. (In the **External Data** window, on the tab for the data recordset, right-click, point to **Data Source**, and then click  **Properties**.)

Methods affected by this property setting include  **[Selection.LinkToData](selection-linktodata-method-visio.md)** , **[Shape.LinkToData](shape-linktodata-method-visio.md)** , and **[Selection.AutomaticLink](selection-automaticlink-method-visio.md)** .

In the UI, when users attempt to link to data shapes that have existing links to data and the setting is  **visLinkReplacePrompt** , Visio responds by opening a dialog box to inform users that their actions will break the existing links and ask if they want to proceed. Because opening a dialog box is not an appropriate response to a method call, when you link shapes by calling any of these methods, Visio treats the setting **visLinkReplacePrompt** as if it were **visLinkReplaceAlways** . That is, these two settings differ in how they affect behavior in the UI, but not programmatic behavior. The default is always to replace existing links when linking is performed programmatically, but to prompt when linking is performed in the UI.

When  **LinkReplaceBehavior** is set to **visLinkReplaceNever** , both of the **LinkToData** methods are disabled and calls to them fail.

The  **LinkReplaceBehavior** setting also affects the default setting of the **Replace Existing Links** check box on the second screen of the **Automatic Link** wizard in the Visio UI (on the **Data** tab, click **Automatically Link**). If  **LinkReplaceBehavior** is set to **visLinkReplaceAlways** or **visLinkReplacePrompt** , this box is selected by default. If the **LinkReplaceBehavior** property is set to **visLinkReplaceNever** , the check box is cleared by default. Users can change the wizard's default behavior by selecting or clearing the check box.

In addition, the  **LinkReplaceBehavior** setting determines how the **Selection.AutomaticLink** method functions. As is the case for the **LinkToData** methods, when **LinkReplaceBehavior** is set to **visLinkReplaceAlways** or **visLinkReplacePrompt** , **AutomaticLink** replaces existing links. And when **LinkReplaceBehavior** is set to **visLinkReplaceNever** , **AutomaticLink** does not replace existing links.

The difference between the  **LinkToData** methods and the **AutomaticLink** method, however, is that for **AutomaticLink** , you can override the **LinkReplaceBehavior** setting by passing either the **visAutoLinkReplaceExistingLinks** or the **visAutoLinkDontReplaceExistingLinks** constant from the **[VisAutoLinkBehaviors](visautolinkbehaviors-enumeration-visio.md)** enumeration to the method as the AutoLinkBehavior parameter.

So, for example, if  **LinkReplaceBehavior** is set to **visLinkReplaceNever** , you can specify that **AutomaticLink** nevertheless replaces existing links by passing it **visAutoLinkReplaceExistingLinks** .


