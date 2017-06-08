---
title: Characters Object (Visio)
keywords: vis_sdr.chm10050
f1_keywords:
- vis_sdr.chm10050
ms.prod: visio
api_name:
- Visio.Characters
ms.assetid: aaff009b-c665-c2ea-8494-e917126d8491
ms.date: 06/08/2017
---


# Characters Object (Visio)

Represents a shape's text with the text fields expanded to the number of characters they display in a drawing window.


## Remarks

To retrieve a  **Characters** object, use the **Characters** property of a **Shape** object.

The default property of a  **Characters** object is **Text**.

The  **Begin** and **End** properties of a **Characters** object determine the range of the shape's text that is represented by the **Characters** object. Initially, the range contains all of the shape's text; you can set the **Begin** and **End** properties to specify a subrange of the text.

After you retrieve a  **Characters** object, you can use its **Text** property to retrieve or set the shape's text. Use the **Copy**, **Cut**, or **Paste** method to copy, cut, or paste the **Characters** object's text to or from the Clipboard. Use the **CharProps** or **ParaProps** property to change the **Characters** object's formatting.

If your Visual Studio solution includes the  **Microsoft.Office.Interop.Visio** reference, this object maps to the following types:


-  **Microsoft.Office.Interop.Visio.IVCharacters**
    

## Events



|**Name**|
|:-----|
|[TextChanged](http://msdn.microsoft.com/library/2387884e-366e-4276-c250-0879fee4cd66%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[AddCustomField](http://msdn.microsoft.com/library/26f3c1b9-36a0-602d-acb2-0a4fcdb7b630%28Office.15%29.aspx)|
|[AddCustomFieldU](http://msdn.microsoft.com/library/f1a5bc23-981d-0be7-92f3-d2ba640751a2%28Office.15%29.aspx)|
|[AddField](http://msdn.microsoft.com/library/1b00cad3-d97a-4bdc-1f8e-cee39d9c836f%28Office.15%29.aspx)|
|[AddFieldEx](http://msdn.microsoft.com/library/14f56159-ed60-e1cf-1c04-b789672b51ec%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/afa21cde-4f1e-cdec-149c-8be7aa88935e%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/08c1e155-335c-0d90-2efa-d079ec14b180%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/b06a2ca3-e0ab-4185-3b46-85fff2dd4cc4%28Office.15%29.aspx)|
|[Paste](http://msdn.microsoft.com/library/e0615a79-b211-643c-15cf-5c6ad8a3cc63%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/88c55936-8dbc-b009-7755-5f5e66484489%28Office.15%29.aspx)|
|[Begin](http://msdn.microsoft.com/library/885adb4d-aca8-b275-806b-34c76a14e7a7%28Office.15%29.aspx)|
|[CharCount](http://msdn.microsoft.com/library/99e780df-b9ee-1083-6efe-cd3e766aa659%28Office.15%29.aspx)|
|[CharProps](http://msdn.microsoft.com/library/7c05633d-9e99-cee3-0d24-bff6d191ef24%28Office.15%29.aspx)|
|[CharPropsRow](http://msdn.microsoft.com/library/55ea568a-7dfc-faed-e4c2-23fa76aac16d%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/50ed7758-208e-15f0-14ac-801db910dabd%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/095cd4fc-1aa1-338a-eb9a-dedb63c2c1ad%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/d685ab44-5db4-65d8-300a-ad40959acdb7%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/61b8fdb4-e00e-b7a5-2f0b-42d46684c626%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/620a254a-9a8d-da0a-1274-305064afdb1c%28Office.15%29.aspx)|
|[FieldCategory](http://msdn.microsoft.com/library/b9c1ecca-ae27-83b8-862d-e8677f8c4c9a%28Office.15%29.aspx)|
|[FieldCode](http://msdn.microsoft.com/library/901e6617-2e4b-6f99-f886-e3c7348a306d%28Office.15%29.aspx)|
|[FieldFormat](http://msdn.microsoft.com/library/298ee3a7-a81e-c10d-e978-ce28ca9408be%28Office.15%29.aspx)|
|[FieldFormula](http://msdn.microsoft.com/library/3bdbf64c-b853-b5bb-6b4f-323d979d3e7e%28Office.15%29.aspx)|
|[FieldFormulaU](http://msdn.microsoft.com/library/83a6f079-bd1a-7512-61f1-0b9fa7c83964%28Office.15%29.aspx)|
|[IsField](http://msdn.microsoft.com/library/329441aa-61ce-177f-061e-a47624a622d2%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/31ffa78e-3232-028b-91a8-636010c9c5b2%28Office.15%29.aspx)|
|[ParaProps](http://msdn.microsoft.com/library/8f71a7ba-3a9e-01b4-1bbe-040fd441a284%28Office.15%29.aspx)|
|[ParaPropsRow](http://msdn.microsoft.com/library/2f87d080-b8a7-d6df-356f-a7cb43453807%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/3cff9c46-6558-322e-8040-7b24218d94a3%28Office.15%29.aspx)|
|[RunBegin](http://msdn.microsoft.com/library/6397f797-c481-e2f0-ec38-61a799762552%28Office.15%29.aspx)|
|[RunEnd](http://msdn.microsoft.com/library/4c9d0f81-8b6d-d5c3-98a1-1d0b39f8193a%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/24565a24-3b95-2a89-1903-ae1759d3d8e2%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/384bd298-e3c4-fed3-d5a0-77f0aa69410a%28Office.15%29.aspx)|
|[TabPropsRow](http://msdn.microsoft.com/library/83002645-df6c-5565-b62a-983960a8a8a3%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/ebfa0548-4150-f6a6-8362-8bd3c2c36f93%28Office.15%29.aspx)|

