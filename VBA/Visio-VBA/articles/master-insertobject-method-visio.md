---
title: Master.InsertObject Method (Visio)
keywords: vis_sdr.chm10716370
f1_keywords:
- vis_sdr.chm10716370
ms.prod: visio
api_name:
- Visio.Master.InsertObject
ms.assetid: 7b663eef-ed40-486b-2b5b-e7c7066c2300
ms.date: 06/08/2017
---


# Master.InsertObject Method (Visio)

Adds a new embedded object or ActiveX control to a page, master, or group.


## Syntax

 _expression_ . **InsertObject**( **_ClassOrProgID_** , **_Flags_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassOrProgID_|Required| **String**|Identifies the type of object or control to create.|
| _Flags_|Required| **Integer**|Flags that influence the operation.|

### Return Value

Shape


## Remarks

 _ClassOrProgID_ is a string that identifies the kind of object or control to create. It can be either the object or control's class ID (GUID) in string form or the object or control's program ID of the handler for the class.




- If  _ClassOrProgID_ is a string representing a class ID, it looks like "{D3E34B21-9D75-101A-8C3D-00AA001A1652}."
    
- If  _ClassOrProgID_ is a string representing a program ID, it looks like "paint.picture" or "forms.combobox.1".
    


See vendor-specific documentation or browse the registry to determine which class IDs and program IDs are associated with objects and controls provided by other applications.

The  _Flags_ argument is a bitmask that can include one of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visInsertIcon**|&;H10|Displays the new shape as an icon.|
| **visInsertDontShow**|&;H1000|Does not execute the new object's show verb.|
If both  **visInsertIcon** and **visInsertDontShow** are specified, the **InsertObject** method fails. If you want to insert an object that is displayed as an icon, you must allow the application to execute the object's show verb.

The  _Flags_ argument can also include one of the following values.



|**Constant **|**Value **|
|:-----|:-----|
| **visInsertAsControl**|&;H2000|
| **visInsertAsEmbed**|&;H4000|
Values in  **visInsertAsControl** and **visInsertAsEmbed** only have an effect if the class identified by _ClassOrProgID_ is identified in the registry as a control that can be inserted. If neither **visInsertAsControl** nor **visInsertAsEmbed** is specified and the object can be either a control or an embedded object, the application inserts it as a control.

In rare cases, Visio 5.0 or later versions may insert a control whereas earlier versions of Visio would have responded to the same call by inserting an embedded object. If a control is inserted, this method places the document in design mode, causing any code executing in the document to halt until the document is returned to run mode.

 **Security** Use caution when you are adding ActiveX controls to your application. ActiveX controls may be designed in such a way that their use could pose a security risk. We recommend that you use controls from trusted sources only. Sign any controls you author.


 **Caution**   Modifying the registry in any manner, whether through the Registry Editor or programmatically, always carries some degree of risk. Incorrect modification can cause serious problems that may require you to reinstall your operating system. It is a good practice to always back up a computer's registry first before modifying it. If you are running Microsoft Windows NT or Microsoft Windows 2000, you should also update your Emergency Repair Disk (ERD).


