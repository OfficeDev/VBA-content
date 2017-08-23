---
title: "Свойство PictureFormat.VerticalScale (издатель)"
keywords: vbapb10.chm3604784
f1_keywords: vbapb10.chm3604784
ms.prod: publisher
api_name: Publisher.PictureFormat.VerticalScale
ms.assetid: ff83d1bc-798b-5b42-7087-9b45f3ff573d
ms.date: 06/08/2017
ms.openlocfilehash: e983132ed8dd309ad694a8847d59dee15c319a4c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatverticalscale-property-publisher"></a>Свойство PictureFormat.VerticalScale (издатель)

Возвращает значение типа **Long** , представляющее горизонтальное рисунок по вертикальной оси. Масштабирование выраженное в процентах (например, равно 200 200% масштабирование). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalScale**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Эффективное разрешение изображения обратно пропорционально масштабирование на печать изображение. Чем больше масштабирование, тем ниже эффективное разрешение. Например предположим, что изображение измерение, 4 4 дюймов сканирования 300 точек на дюйм. Если этот рисунок масштабируется 2 дюйма с 2 дюйма, эффективное разрешение — 600 точек на дюйм.

Используйте свойство **[EffectiveResolution](pictureformat-effectiveresolution-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** для определения разрешения, в котором этот рисунок или объект OLE печатает в указанный документ.


## <a name="example"></a>Пример

В следующем примере выводится свойства выбранного изображения для каждого изображения в активной публикации.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Resolution in Publication: " &; .EffectiveResolution &; " dpi" 
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Height in publication: " &; .Height &; " points" 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "Width in publication: " &; .Width &; " points" 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 
 

```


