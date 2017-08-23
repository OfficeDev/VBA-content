---
title: "Свойство PictureFormat.Contrast (издатель)"
keywords: vbapb10.chm3604738
f1_keywords: vbapb10.chm3604738
ms.prod: publisher
api_name: Publisher.PictureFormat.Contrast
ms.assetid: f081b7c8-50cc-772b-f3b0-27c215cfebac
ms.date: 06/08/2017
ms.openlocfilehash: 1e87fa0a3d37c3d848e0eead0a4dfd0ff6c58803
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcontrast-property-publisher"></a>Свойство PictureFormat.Contrast (издатель)

Возвращает или задает **единого** , указывающее, контрастности для указанного изображения или объекта OLE. Значение для этого свойства должна быть число от 0,0 (как минимум контрастности) до 1.0 (наивысшего контрастности). Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Контрастности**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Один


## <a name="remarks"></a>Заметки

Используйте метод **[IncrementContrast](pictureformat-incrementcontrast-method-publisher.md)** для постепенного настройки контрастности из его текущего уровня.


## <a name="example"></a>Пример

В этом примере задается контрастности первую фигуру в активной публикации. Фигура должен быть изображения или объекта OLE.


```vb
ActiveDocument.Pages(1).Shapes(1).PictureFormat _ 
 .Contrast = 0.8
```


