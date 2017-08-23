---
title: "Метод Shapes.AddEmptyPictureFrame (издатель)"
keywords: vbapb10.chm2162757
f1_keywords: vbapb10.chm2162757
ms.prod: publisher
api_name: Publisher.Shapes.AddEmptyPictureFrame
ms.assetid: e473dea8-6d94-e9e4-ddb6-27c1fc8930e8
ms.date: 06/08/2017
ms.openlocfilehash: ed9e343608dfd9c80ee07daee04ad3f8d457b8b8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddemptypictureframe-method-publisher"></a>Метод Shapes.AddEmptyPictureFrame (издатель)

Возвращает объект **фигуры** , представляющий пустой рамки вставлен по указанным координатам.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddEmptyPictureFrame** ( **_Слева_**, **_сверху_**, **_Ширина_**, **_Высота_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Слева|Обязательное свойство.| **Variant**|Положение левого края фигуры, представляющее рисунок.|
|Вверх|Обязательное свойство.| **Variant**|Положение верхнего края фигуры, представляющее рисунок.|
|Width|Необязательный| **Variant**|Ширина формы, представляющее рисунок. По умолчанию используется значение -1, что означает, что ширину фигуры автоматически устанавливается значение 72 точки, если не указан параметр.|
|Height|Необязательный| **Variant**|Высота формы, представляющее рисунок. По умолчанию используется значение -1, что означает, что высоту фигуры автоматически устанавливается значение 54 точек, если не указан параметр.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

**Слева**, **в начало**, **ширину**и **высоту** аргументы числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «1,5 в»).

Пустая рамка имеет по умолчанию подсказка «Пустая рамка рисунка». Это изменяется на «Рисунок» при выборе изображения для **фигуры**.


## <a name="example"></a>Пример

В этом примере помещает пустой рамки в центре первой страницы публикации и разворачивает его на 45 градусов. Свойство **AlternativeText** имеет значение «Рисунок 1» для веб-сайта.


```vb
Dim shpPlaceholder As Shape 
 
Set shpPlaceholder = _ 
 ActiveDocument.Pages(1).Shapes.AddEmptyPictureFrame( _ 
 230, 320, 150, 150) 
 
With shpPlaceholder 
 .AlternativeText = "Picture Placeholder 1" 
 .Rotation = 45 
End With 
 

```


