---
title: "Свойство PictureFormat.Brightness (издатель)"
keywords: vbapb10.chm3604736
f1_keywords: vbapb10.chm3604736
ms.prod: publisher
api_name: Publisher.PictureFormat.Brightness
ms.assetid: bed1cd25-faee-6fb9-4bb3-5bdaf148b62e
ms.date: 06/08/2017
ms.openlocfilehash: 38272e935d6e83a3054c58176ee2df55fca5f568
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatbrightness-property-publisher"></a>Свойство PictureFormat.Brightness (издатель)

Возвращает или задает **единого** , указывающее, яркость указанного изображения или объекта OLE. Значение этого свойства должно быть числом между 0.0 (dimmest) 1.0 (ярким). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Яркость**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[IncrementBrightness](pictureformat-incrementbrightness-method-publisher.md)** для постепенного яркость из его текущего уровня.


## <a name="example"></a>Пример

В этом примере задается яркость для первой фигуры в активной публикации. Фигура должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .Brightness = 0.3
```


