---
title: "Свойство PictureFormat.HorizontalPictureLocking (издатель)"
keywords: vbapb10.chm3604752
f1_keywords: vbapb10.chm3604752
ms.prod: publisher
api_name: Publisher.PictureFormat.HorizontalPictureLocking
ms.assetid: 9a8cb8ec-24d1-4a21-d662-bcdfd26821df
ms.date: 06/08/2017
ms.openlocfilehash: 0d980d5946bfafbde31ed324051aad4edb182f0b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformathorizontalpicturelocking-property-publisher"></a>Свойство PictureFormat.HorizontalPictureLocking (издатель)

Возвращает или задает значение, указывающее, где отображаются вставленных новых изображений при использовании указанного кадра константы **PbHorizontalPictureLocking** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HorizontalPictureLocking**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbHorizontalPictureLocking


## <a name="remarks"></a>Заметки

Значение свойства **HorizontalPictureLocking** может иметь одно из **[PbHorizontalPictureLocking](pbhorizontalpicturelocking-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

Следующий пример блокирует указанный рисунок в верхний левый угол рамка рисунка. Фигура одно на странице один из активных публикации должен быть рамка рисунка для работы этого примера.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .HorizontalPictureLocking = pbHorizontalLockingLeft 
 .VerticalPictureLocking = pbVerticalLockingTop 
End With
```


