---
title: "Метод Shapes.AddWebNavigationBar (издатель)"
keywords: vbapb10.chm2162736
f1_keywords: vbapb10.chm2162736
ms.prod: publisher
api_name: Publisher.Shapes.AddWebNavigationBar
ms.assetid: 26e9622c-ea28-b28b-9904-b3a3ccc9341b
ms.date: 06/08/2017
ms.openlocfilehash: 1d0c9ac6f09e6754a6afcec6fbdd1b65819ff51a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddwebnavigationbar-method-publisher"></a>Метод Shapes.AddWebNavigationBar (издатель)

Добавляет объект **фигуры** типа **pbWebNavigationBar** текущей страницы публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddWebNavigationBar** ( **_Имя_**, **_слева_** **_в начало_**, **_Ширина_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя|Обязательное свойство.| **String**|Имя объекта **WebNavigationBarSet** для добавления указанного **фигуры**.|
|Слева|Обязательное свойство.| **Variant**|Задайте положение левого края фигуры, представляющий панель навигации.|
|Вверх|Обязательное свойство.| **Variant**|Задайте положение верхнего края фигуры, представляющий панель навигации.|
|Width|Необязательный| **Variant**|Задать ширину фигуры, представляющий панель навигации.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Метод **AddWebNavigationBar** создает набор панель навигации Web. Добавление существующего набора из коллекции **WebNavigationBarSets** . Передайте имя Web панель навигации задайте в качестве имени параметра.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **AddWebNavigationBar** для добавления объекта **WebNavigationBarSet** в активный документ.


```vb
Public Sub AddWebNavigationBarSet_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 ThisDocument.WebNavigationBarSets.AddSet ("NavBar") 
 Set pubShape = ThisDocument.Pages(1).Shapes.AddWebNavigationBar("NavBar", 10, 25) 
 
End Sub
```


