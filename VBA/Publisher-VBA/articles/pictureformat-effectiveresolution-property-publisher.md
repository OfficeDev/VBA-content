---
title: "Свойство PictureFormat.EffectiveResolution (издатель)"
keywords: vbapb10.chm3604755
f1_keywords: vbapb10.chm3604755
ms.prod: publisher
api_name: Publisher.PictureFormat.EffectiveResolution
ms.assetid: 33e5323f-5e10-b2ed-62eb-03ecbbb1e893
ms.date: 06/08/2017
ms.openlocfilehash: e379b3229baf262c31dd998e5622df85a17c3d48
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformateffectiveresolution-property-publisher"></a>Свойство PictureFormat.EffectiveResolution (издатель)

Возвращает значение типа **Long** , представляющий, точек на дюйм (т/д), эффективное разрешение рисунка. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EffectiveResolution**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Эффективное разрешение изображения обратно пропорционально масштабирование на печать изображение. Чем больше масштабирование, тем ниже эффективное разрешение. Например предположим, что изображение измерение, 4 4 дюймов сканирования 300 точек на дюйм. Если этот рисунок масштабируется 2 дюйма с 2 дюйма, эффективное разрешение — 600 точек на дюйм.

Используйте свойство **[OriginalResolution](pictureformat-originalresolution-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** для определения разрешения связанных рисунков или объекты OLE. Использование свойства **[HorizontalScale](pictureformat-horizontalscale-property-publisher.md)** и **[VerticalScale](pictureformat-verticalscale-property-publisher.md)** для определения масштабирование изображения.


## <a name="example"></a>Пример

Следующий пример возвращает список рисунков с фактическим разрешением падает ниже указанного порогового значения (100 точек на дюйм) в активной публикации.


```vb
Sub ListLowResolutionPictures() 
 Dim pgLoop As Page 
 Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 If .EffectiveResolution < 100 Then 
 Debug.Print .Filename 
 Debug.Print "Page " &; pgLoop.PageNumber 
 Debug.Print "Resolution in publication: " &; .EffectiveResolution 
 End If 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


