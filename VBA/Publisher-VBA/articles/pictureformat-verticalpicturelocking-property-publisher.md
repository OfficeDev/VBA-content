---
title: "Свойство PictureFormat.VerticalPictureLocking (издатель)"
keywords: vbapb10.chm3604745
f1_keywords: vbapb10.chm3604745
ms.prod: publisher
api_name: Publisher.PictureFormat.VerticalPictureLocking
ms.assetid: 0575d733-b515-2256-7136-6ec07532ab67
ms.date: 06/08/2017
ms.openlocfilehash: 3f4cc573d2569b55521dd2c1f9522d5cf2032c4a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatverticalpicturelocking-property-publisher"></a>Свойство PictureFormat.VerticalPictureLocking (издатель)

Возвращает или задает значение, указывающее, где отображаются вставленных новых изображений при использовании указанного кадра константы **PbVerticalPictureLocking** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **VerticalPictureLocking**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbVerticalPictureLocking


## <a name="remarks"></a>Заметки

Значение свойства **Вертикали PictureLocking** может быть одной из констант **PbVerticalPictureLocking** объявлена в библиотеке типов, Microsoft Publisher и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **pbVerticalLockingBottom**|Новые изображения, вставляются нижнему краю фрейма.|
| **pbVerticalLockingNone**|Новые изображения вставляется в центре между верхнего и нижнего края кадра.|
| **pbVerticalLockingStretch**|Новые изображения по вертикали расширяются полный высота кадра.|
| **pbVerticalLockingTop**|Новые изображения, вставляются по верхней границе фрейма.|

## <a name="example"></a>Пример

Следующий пример блокирует указанный рисунок в верхний левый угол рамка рисунка. Фигура одно на странице один из активных публикации должен быть рамка рисунка для работы этого примера.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .HorizontalPictureLocking = pbHorizontalLockingLeft 
 .VerticalPictureLocking = pbVerticalLockingTop 
End With
```


