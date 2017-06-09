---
title: Cell Object (Visio)
keywords: vis_sdr.chm10045
f1_keywords:
- vis_sdr.chm10045
ms.prod: visio
api_name:
- Visio.Cell
ms.assetid: 06ac28a6-5749-6c70-94bf-c721e217f375
ms.date: 06/08/2017
---


# Cell Object (Visio)

Holds a formula that evaluates to some value.


## Remarks

The default property of a  **Cell** object is **ResultIU**.

You can get or set a cell's formula or value. A cell belongs to a  **Shape**, **Style**, or **Row** object and represents a property of the shape, style, or row. For example, the height of a shape equals the value of the shape's Height cell.

A program can control a shape's appearance and behavior by working with the formulas in the shape's cells. You can visually inspect most of a shape's cells by opening the shape's ShapeSheet window. Use the  **Cells** or **CellsSRC** property of a **Shape** object to retrieve a **Cell** object. To retrieve a cell in a style, use the **Cells** property of a **Style** object.


## Events



|**Name**|
|:-----|
|[CellChanged](http://msdn.microsoft.com/library/f39c2a33-bff9-ee67-1bfe-618f5d702c8b%28Office.15%29.aspx)|
|[FormulaChanged](http://msdn.microsoft.com/library/7f612470-ea40-1b7e-7334-825a124a96f3%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[GlueTo](http://msdn.microsoft.com/library/dc88ecf1-d7c2-994e-8b49-e4bfddef4472%28Office.15%29.aspx)|
|[GlueToPos](http://msdn.microsoft.com/library/9f9e10f2-030f-f7ad-be04-ea2804c20cb4%28Office.15%29.aspx)|
|[Trigger](http://msdn.microsoft.com/library/aea545d3-5e5d-2206-c0fe-c062bc4e6be8%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/ec2bd6fb-5c24-acf2-7324-e8db42d903a9%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/4850bc13-96dc-bb68-8c92-476fc430d969%28Office.15%29.aspx)|
|[ContainingMasterID](http://msdn.microsoft.com/library/1daba8ed-69cd-2c80-8534-ba9fc4956292%28Office.15%29.aspx)|
|[ContainingPageID](http://msdn.microsoft.com/library/0d4c97cc-d84e-c13e-759b-8805114d191e%28Office.15%29.aspx)|
|[ContainingRow](http://msdn.microsoft.com/library/ebe3f83c-6c97-c652-70d1-fb1197873ffb%28Office.15%29.aspx)|
|[Dependents](http://msdn.microsoft.com/library/99a1502b-c847-6836-2470-178b595345f9%28Office.15%29.aspx)|
|[Document](http://msdn.microsoft.com/library/170f93ef-d60f-c683-a840-f2168479a80d%28Office.15%29.aspx)|
|[Error](http://msdn.microsoft.com/library/8c2966b7-f734-cb3a-7bc0-24c2d9575125%28Office.15%29.aspx)|
|[EventList](http://msdn.microsoft.com/library/d88186f6-ecf6-c645-5250-46e07862a745%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/36636047-9ee7-d461-92fb-0b36853e3201%28Office.15%29.aspx)|
|[FormulaForce](http://msdn.microsoft.com/library/bce2a3c8-eaac-42dc-3a7f-c4838ed6518b%28Office.15%29.aspx)|
|[FormulaForceU](http://msdn.microsoft.com/library/386003e3-b9e9-4c35-ac14-55bdb8da4375%28Office.15%29.aspx)|
|[FormulaU](http://msdn.microsoft.com/library/931490f6-938c-f783-eb2f-a67505187c90%28Office.15%29.aspx)|
|[InheritedFormulaSource](http://msdn.microsoft.com/library/62aedef3-06b1-2fc3-5fd2-03f77668548f%28Office.15%29.aspx)|
|[InheritedValueSource](http://msdn.microsoft.com/library/1ffa8293-80a9-a43b-c6e1-b90cb2648efa%28Office.15%29.aspx)|
|[IsConstant](http://msdn.microsoft.com/library/ed17029d-9044-d6fe-aac0-81fd8ac74b56%28Office.15%29.aspx)|
|[IsInherited](http://msdn.microsoft.com/library/e68ef657-64dc-2e8e-d21f-d8ff5566a12d%28Office.15%29.aspx)|
|[LocalName](http://msdn.microsoft.com/library/596bf196-6bbc-32f0-e508-03cdf4969a7f%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/293cfa05-7eb8-98d2-0080-378df17a4408%28Office.15%29.aspx)|
|[ObjectType](http://msdn.microsoft.com/library/9abf9d16-e996-2283-5caf-0767b9fdd0a4%28Office.15%29.aspx)|
|[PersistsEvents](http://msdn.microsoft.com/library/08e3095e-37ef-65f1-7109-b2f9deeeec14%28Office.15%29.aspx)|
|[Precedents](http://msdn.microsoft.com/library/4461b45a-6fd6-4376-f8b2-4d8a9597111a%28Office.15%29.aspx)|
|[Result](http://msdn.microsoft.com/library/5d97f8e7-0bb4-7334-8cf0-7fb3860fbc2b%28Office.15%29.aspx)|
|[ResultForce](http://msdn.microsoft.com/library/96579953-05f2-edf5-02d6-54ef0e632215%28Office.15%29.aspx)|
|[ResultFromInt](http://msdn.microsoft.com/library/1fb4b39b-b868-64b1-1952-405045a11d6f%28Office.15%29.aspx)|
|[ResultFromIntForce](http://msdn.microsoft.com/library/e22b2479-a55f-c08b-4d2b-18f8225900fa%28Office.15%29.aspx)|
|[ResultInt](http://msdn.microsoft.com/library/f3e2ef7d-cde1-a0d4-3d02-f5bf329cd0c3%28Office.15%29.aspx)|
|[ResultIU](http://msdn.microsoft.com/library/4d752d78-e112-bb45-08c7-5411d7d79beb%28Office.15%29.aspx)|
|[ResultIUForce](http://msdn.microsoft.com/library/ae26cf67-5f4c-6431-82ad-0866eac0fabd%28Office.15%29.aspx)|
|[ResultStr](http://msdn.microsoft.com/library/f5d1236b-2596-298c-1ad4-6e19f5c32ef4%28Office.15%29.aspx)|
|[ResultStrU](http://msdn.microsoft.com/library/2a2fc8c9-eb2c-6c49-9af6-abc120bbd610%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/b31b981d-8034-db03-b7db-06eb98ac744b%28Office.15%29.aspx)|
|[RowName](http://msdn.microsoft.com/library/4f5f57f9-c147-5991-c3f0-2caad2993d77%28Office.15%29.aspx)|
|[RowNameU](http://msdn.microsoft.com/library/3c73ed3d-851f-faf4-fab0-76d6602da82b%28Office.15%29.aspx)|
|[Section](http://msdn.microsoft.com/library/aab8e98c-e28b-033e-1c29-852f5ad2861f%28Office.15%29.aspx)|
|[Shape](http://msdn.microsoft.com/library/4929ea4e-6498-8ddc-1c38-1276043aaa4e%28Office.15%29.aspx)|
|[Stat](http://msdn.microsoft.com/library/9421a8f1-8cc1-2e29-b145-958908a3efe9%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/12eec8c7-706a-488e-ad3a-326c9f628f5c%28Office.15%29.aspx)|
|[Units](http://msdn.microsoft.com/library/075cfda9-8b7a-550b-cf72-b8044c3d461a%28Office.15%29.aspx)|

