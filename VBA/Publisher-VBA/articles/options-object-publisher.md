---
title: Options Object (Publisher)
keywords: vbapb10.chm1114111
f1_keywords:
- vbapb10.chm1114111
ms.prod: publisher
api_name:
- Publisher.Options
ms.assetid: 2554cd33-9d94-2622-6fab-19ca33d5a561
ms.date: 06/08/2017
---


# Options Object (Publisher)

Represents application and publication options in Microsoft Publisher. Many of the properties for the  **Options** object correspond to items in the **Options** dialog box ( **Tools** menu).


## Example

Use the  **[Options](http://msdn.microsoft.com/library/999f208a-02e6-49fb-c9a0-42aa97c5e37e%28Office.15%29.aspx)** property to return the **Options** object. The following example sets four application options for Publisher.


```
Sub SetSpecialOptions() 
 With Options 
 .AllowBackgroundSave = True 
 .DragAndDropText = True 
 .AutoHyphenate = True 
 .MeasurementUnit = pbUnitInch 
 End With 
End Sub
```


## Methods



|**Name**|
|:-----|
|[ResetTips](http://msdn.microsoft.com/library/a119aacc-ba19-f430-e8af-6d84c438ec25%28Office.15%29.aspx)|
|[ResetWizardSynchronizing](http://msdn.microsoft.com/library/1027a113-45aa-b722-b625-a6bb7bbcc3e6%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddHebDoubleQuote](http://msdn.microsoft.com/library/9c71b52e-0273-7ca9-1f50-5beed65c2e73%28Office.15%29.aspx)|
|[AllowBackgroundSave](http://msdn.microsoft.com/library/5bddfb2d-7fb7-99db-43ea-c6ee53e1d0b3%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/06336d0e-18c8-f364-7911-1749d125d638%28Office.15%29.aspx)|
|[AutoFormatWord](http://msdn.microsoft.com/library/b0466bd7-f0a1-44a8-480f-5d046e24e759%28Office.15%29.aspx)|
|[AutoHyphenate](http://msdn.microsoft.com/library/821d0540-80ec-9f9d-777e-4d2596baf7d7%28Office.15%29.aspx)|
|[AutoKeyboardSwitching](http://msdn.microsoft.com/library/05f22aa6-332d-e033-ab9d-550eb08f1018%28Office.15%29.aspx)|
|[AutoSelectWord](http://msdn.microsoft.com/library/2b36f0d2-3260-aa3d-13b2-ae08b8d631d1%28Office.15%29.aspx)|
|[DefaultPubDirection](http://msdn.microsoft.com/library/628352c1-040f-9ab1-d0f1-308b2c26679c%28Office.15%29.aspx)|
|[DefaultTextFlowDirection](http://msdn.microsoft.com/library/7c17768a-cd9c-704d-fa27-f0dfd7648054%28Office.15%29.aspx)|
|[DisplayStatusBar](http://msdn.microsoft.com/library/335b2f1e-03ff-fd90-5ec2-27d5219b27e7%28Office.15%29.aspx)|
|[DragAndDropText](http://msdn.microsoft.com/library/55fb68e8-4ddc-6866-00d8-bdd6a1e25ec3%28Office.15%29.aspx)|
|[HyphenationZone](http://msdn.microsoft.com/library/ed0e90de-4a2a-3c8a-27f1-e8c7c1f0e174%28Office.15%29.aspx)|
|[MeasurementUnit](http://msdn.microsoft.com/library/49221e4e-c84a-6706-8f9a-3853283ebb18%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/96b43655-699c-96cc-bfc9-14199619b699%28Office.15%29.aspx)|
|[PathForPictures](http://msdn.microsoft.com/library/e66c8c86-f049-0f32-0a0d-60fd37470708%28Office.15%29.aspx)|
|[PathForPublications](http://msdn.microsoft.com/library/d33d5eab-eb52-b533-8968-31ddb5e12d99%28Office.15%29.aspx)|
|[SaveAutoRecoverInfo](http://msdn.microsoft.com/library/1cbb7960-8995-37f4-5989-01b97152269f%28Office.15%29.aspx)|
|[SaveAutoRecoverInfoInterval](http://msdn.microsoft.com/library/3d6a6c4f-7e2b-18ff-67a4-20dee4fbcf5b%28Office.15%29.aspx)|
|[SequenceCheck](http://msdn.microsoft.com/library/a2801af8-5c89-9256-80a6-d9dac17b6066%28Office.15%29.aspx)|
|[ShowBasicColors](http://msdn.microsoft.com/library/d04504fa-5627-b66b-bd6e-30556155632c%28Office.15%29.aspx)|
|[ShowScreenTipsOnObjects](http://msdn.microsoft.com/library/b5503200-31fd-72ac-de28-ace55a7123b3%28Office.15%29.aspx)|
|[ShowTipPages](http://msdn.microsoft.com/library/44f91cf1-68e3-0755-3114-5dc41a2e4eba%28Office.15%29.aspx)|
|[TypeNReplace](http://msdn.microsoft.com/library/0eb378d2-3554-6a46-8b6b-4a990b4638db%28Office.15%29.aspx)|
|[UseCatalogAtStartup](http://msdn.microsoft.com/library/7b0cfce9-92f1-5491-c550-421d1c848e0f%28Office.15%29.aspx)|
|[UseWizardForBlankPublication](http://msdn.microsoft.com/library/c8afb883-03db-0ec4-1a7a-ebac697fc72f%28Office.15%29.aspx)|

