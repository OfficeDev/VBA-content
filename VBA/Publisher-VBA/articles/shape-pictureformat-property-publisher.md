---
title: "Свойство Shape.PictureFormat (издатель)"
keywords: vbapb10.chm2228295
f1_keywords: vbapb10.chm2228295
ms.prod: publisher
api_name: Publisher.Shape.PictureFormat
ms.assetid: 2a812ba3-18e4-fc42-6d07-535511a79650
ms.date: 06/08/2017
ms.openlocfilehash: a2b5dd9d4934209f1e99e6911616abef348bdf8e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapepictureformat-property-publisher"></a>Свойство Shape.PictureFormat (издатель)

Возвращает объект **[PictureFormat](pictureformat-object-publisher.md)** , который содержит изображение свойства для указанного объекта форматирования. Применяется к **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** объектов, которые представляют изображения или объекты OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PictureFormat**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="example"></a>Пример

В этом примере задается яркости и контрастности для всех рисунков на первой странице active публикации.


```vb
Sub FixPictureContrastBrightness() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = pbPicture Then 
 With shp.PictureFormat 
 .Brightness = 0.6 
 .Contrast = 0.6 
 End With 
 End If 
 Next shp 
End Sub
```


