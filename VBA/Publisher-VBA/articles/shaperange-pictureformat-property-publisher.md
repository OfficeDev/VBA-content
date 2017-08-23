---
title: "Свойство ShapeRange.PictureFormat (издатель)"
keywords: vbapb10.chm2293831
f1_keywords: vbapb10.chm2293831
ms.prod: publisher
api_name: Publisher.ShapeRange.PictureFormat
ms.assetid: 3d693c6b-b76b-0fe1-e7df-63fb08782f6f
ms.date: 06/08/2017
ms.openlocfilehash: 9c6118ff72a49ac0a8570b1e75743d2a06576ab6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangepictureformat-property-publisher"></a>Свойство ShapeRange.PictureFormat (издатель)

Возвращает объект **[PictureFormat](pictureformat-object-publisher.md)** , который содержит изображение свойства для указанного объекта форматирования. Применяется к **[фигуры](shape-object-publisher.md)** или **[ShapeRange](shaperange-object-publisher.md)** объектов, которые представляют изображения или объекты OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PictureFormat**

 переменная _expression_A, представляющий объект **ShapeRange** .


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


