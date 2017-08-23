---
title: "Свойство View.Zoom (издатель)"
keywords: vbapb10.chm327684
f1_keywords: vbapb10.chm327684
ms.prod: publisher
api_name: Publisher.View.Zoom
ms.assetid: 31727291-740b-4e77-9c6b-9f19523488cb
ms.date: 06/08/2017
ms.openlocfilehash: 8f71b13ea33f17314e2b62cf728c48de60019d53
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="viewzoom-property-publisher"></a>Свойство View.Zoom (издатель)

Возвращает или задает константа **PbZoom** или в диапазоне от 10 до 400, указывая значения масштабирования указанное представление. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Показать**

 переменная _expression_A, представляющий объект **View** .


### <a name="return-value"></a>Возвращаемое значение

PbZoom


## <a name="remarks"></a>Заметки

**Отобразить** значение свойства может быть одной из констант **PbZoom** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbZoomFitSelection**| Изменяет размер страницы представления на размер текущего выделенного фрагмента.|
| **pbZoomPageWidth**|Изменяет размер страницы представления ширину публикации. |
| **pbZoomWholePage**| Изменяет размер страницы представления на размер всей страницы.|

## <a name="example"></a>Пример

В следующем примере задается масштаб для активной публикации, чтобы образом подходит для всей страницы на экране.


```vb
ActiveDocument.ActiveView.Zoom = pbZoomWholePage
```


