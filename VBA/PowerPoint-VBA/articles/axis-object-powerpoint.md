---
title: Axis Object (PowerPoint)
keywords: vbapp10.chm682000
f1_keywords:
- vbapp10.chm682000
ms.prod: powerpoint
api_name:
- PowerPoint.Axis
ms.assetid: 38d5e006-ac32-7bdb-f9f0-e8a858dcbf49
ms.date: 06/08/2017
---


# Axis Object (PowerPoint)

Represents a single axis in a chart.


## Remarks

The  **Axis** object is a member of the **[Axes](http://msdn.microsoft.com/library/71f1e1fc-7086-a84e-1e05-6fa50597b49b%28Office.15%29.aspx)** collection.

Use  **Axes** ( _Type_, _AxisGroup_ ) where _Type_ is the axis type and _AxisGroup_ is the axis group to return a single **Axis** object. _Type_ can be one of the following **[XlAxisType](http://msdn.microsoft.com/library/6eb891d5-3b69-e0a4-90e5-0b21afb1eeaa%28Office.15%29.aspx)** constants: **xlCategory**, **xlSeries**, or **xlValue**. _AxisGroup_ can be one of the following **[XlAxisGroup](http://msdn.microsoft.com/library/775041e9-c965-a9b6-b5fb-cdebe4fb71c0%28Office.15%29.aspx)** constants: **xlPrimary** or **xlSecondary**. For more information, see the **[Axes](http://msdn.microsoft.com/library/6f740a9e-2baa-5a84-ea51-6a39452e227e%28Office.15%29.aspx)** method.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the category axis title text for the first chart in the active document.




```
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        With .Chart.Axes(xlCategory)

            .HasTitle = True

            .AxisTitle.Caption = "1994"

        End With

    End If

End With
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/be589a1e-0484-dffc-f514-fc93c377f9c2%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/1bc059fa-f8b5-f3be-64e2-462dc9cee175%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/eec63378-6282-bf8e-b421-fca7a8b4e95c%28Office.15%29.aspx)|
|[AxisBetweenCategories](http://msdn.microsoft.com/library/8e0e0e80-58b9-005f-c719-ad45b491f9a9%28Office.15%29.aspx)|
|[AxisGroup](http://msdn.microsoft.com/library/19261289-1677-cbd2-70a4-4109bed4b554%28Office.15%29.aspx)|
|[AxisTitle](http://msdn.microsoft.com/library/c1063cf8-3aa2-ea39-ea2d-33a7c63b77d4%28Office.15%29.aspx)|
|[BaseUnit](http://msdn.microsoft.com/library/a53e90c5-5048-8e93-57b2-024d64d2ff73%28Office.15%29.aspx)|
|[BaseUnitIsAuto](http://msdn.microsoft.com/library/3cc90d1a-a87f-ac57-b2a2-bf3ccc964a8e%28Office.15%29.aspx)|
|[Border](http://msdn.microsoft.com/library/fee770aa-879b-17ab-0906-1b0c1faa8a2b%28Office.15%29.aspx)|
|[CategoryNames](http://msdn.microsoft.com/library/f076ad9f-819b-4ced-a967-2fbda72fdfe8%28Office.15%29.aspx)|
|[CategoryType](http://msdn.microsoft.com/library/bbcb485d-9464-33c8-ca9b-e3463bc9e884%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/ae5c24b1-3bf4-e285-7402-12f2a4727e2a%28Office.15%29.aspx)|
|[Crosses](http://msdn.microsoft.com/library/93390bc6-8d94-4bf3-257e-c20fce2a2c62%28Office.15%29.aspx)|
|[CrossesAt](http://msdn.microsoft.com/library/ccc6438d-fb72-7674-0994-bf5d9cecf58d%28Office.15%29.aspx)|
|[DisplayUnit](http://msdn.microsoft.com/library/6545b191-ef58-49d5-2df3-04d0d0d06476%28Office.15%29.aspx)|
|[DisplayUnitCustom](http://msdn.microsoft.com/library/bfee899d-27fd-ca15-9af7-04702ae3da52%28Office.15%29.aspx)|
|[DisplayUnitLabel](http://msdn.microsoft.com/library/75b01ce4-8edd-bbaa-d0fb-2d36c96b4da6%28Office.15%29.aspx)|
|[Format](http://msdn.microsoft.com/library/c00a6cbf-d2d1-3718-2cd6-d61abeed40d3%28Office.15%29.aspx)|
|[HasDisplayUnitLabel](http://msdn.microsoft.com/library/adbbbb89-55af-12f5-ec67-1e88424f3d81%28Office.15%29.aspx)|
|[HasMajorGridlines](http://msdn.microsoft.com/library/a8d5a060-ce84-8ca5-a42c-4a52d09a1e50%28Office.15%29.aspx)|
|[HasMinorGridlines](http://msdn.microsoft.com/library/4ee1c716-296b-eeaf-8d14-bcb6e0919611%28Office.15%29.aspx)|
|[HasTitle](http://msdn.microsoft.com/library/04f9e10a-f323-a905-e09c-e9bb3222a80c%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/d5dc2035-fa09-4a77-2cb4-dc44659efd9e%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/d01f11d2-69e0-1415-6418-0682e40fc6b5%28Office.15%29.aspx)|
|[LogBase](http://msdn.microsoft.com/library/e071420c-6940-4ba8-28b3-d19fe1d844c3%28Office.15%29.aspx)|
|[MajorGridlines](http://msdn.microsoft.com/library/d0ec2384-8503-0198-388c-c74231137bf0%28Office.15%29.aspx)|
|[MajorTickMark](http://msdn.microsoft.com/library/82397f1c-8a0d-44dd-a9f3-3426fee03f1d%28Office.15%29.aspx)|
|[MajorUnit](http://msdn.microsoft.com/library/5f88f369-e999-b947-c47f-5413e349d192%28Office.15%29.aspx)|
|[MajorUnitIsAuto](http://msdn.microsoft.com/library/ffea2f83-1a5e-7ae1-f866-ae52a4d49567%28Office.15%29.aspx)|
|[MajorUnitScale](http://msdn.microsoft.com/library/42fe928b-de99-02d5-b56e-1e735ba5f0da%28Office.15%29.aspx)|
|[MaximumScale](http://msdn.microsoft.com/library/cb0588ce-0685-77ac-da06-75a913f90e41%28Office.15%29.aspx)|
|[MaximumScaleIsAuto](http://msdn.microsoft.com/library/f25fd6a9-4ca7-2f06-3db4-35002f1c91ae%28Office.15%29.aspx)|
|[MinimumScale](http://msdn.microsoft.com/library/90cfaa99-0ccf-2fa8-c6b0-54d1440b6677%28Office.15%29.aspx)|
|[MinimumScaleIsAuto](http://msdn.microsoft.com/library/7ec5b07d-3683-e45b-ca39-d67ce959edfc%28Office.15%29.aspx)|
|[MinorGridlines](http://msdn.microsoft.com/library/f9e1168d-af71-6876-a289-a9e8d1db38cb%28Office.15%29.aspx)|
|[MinorTickMark](http://msdn.microsoft.com/library/2486a649-7006-388f-1b52-379b44f3f80d%28Office.15%29.aspx)|
|[MinorUnit](http://msdn.microsoft.com/library/ff4b4a7b-25b3-974c-dee1-b81f0ec15634%28Office.15%29.aspx)|
|[MinorUnitIsAuto](http://msdn.microsoft.com/library/18dff25c-59a3-e2c8-2997-6239b1ae87bf%28Office.15%29.aspx)|
|[MinorUnitScale](http://msdn.microsoft.com/library/15ce78c6-b054-afea-bd6c-6a40db7f93aa%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/14409079-3cd4-7581-295a-adbd093dbdad%28Office.15%29.aspx)|
|[ReversePlotOrder](http://msdn.microsoft.com/library/630d989b-1f9b-5258-d0be-479f362d2c66%28Office.15%29.aspx)|
|[ScaleType](http://msdn.microsoft.com/library/baf40097-28a4-c2ec-fea9-2ce971f72ed5%28Office.15%29.aspx)|
|[TickLabelPosition](http://msdn.microsoft.com/library/439b3da0-37d1-1fd8-b810-66accac03001%28Office.15%29.aspx)|
|[TickLabels](http://msdn.microsoft.com/library/80e39b06-b01d-f817-5357-e6abbbc28e1c%28Office.15%29.aspx)|
|[TickLabelSpacing](http://msdn.microsoft.com/library/9a6694cb-bb6c-fc5d-a2a3-656327121581%28Office.15%29.aspx)|
|[TickLabelSpacingIsAuto](http://msdn.microsoft.com/library/f0c644a4-2842-6468-9047-239f891dd0b2%28Office.15%29.aspx)|
|[TickMarkSpacing](http://msdn.microsoft.com/library/85c37d23-b91a-b390-4475-a4afa21d1566%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/058723d8-ca0f-3139-b5cc-6f31fe9ff8f9%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/f0bf5b51-fc39-060e-6030-657e7b22fa59%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/a9f618a4-36c4-1e05-8c0c-68edc7608417%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
