---
title: "Метод Shape.MoveToPage (издатель)"
keywords: vbapb10.chm2228376
f1_keywords: vbapb10.chm2228376
ms.prod: publisher
api_name: Publisher.Shape.MoveToPage
ms.assetid: 1893035f-6739-7480-6ba0-2ca6a42355fa
ms.date: 06/08/2017
ms.openlocfilehash: 5cdb7940eb0ff9d4dd956e4d2b7f5f42f15a706c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapemovetopage-method-publisher"></a>Метод Shape.MoveToPage (издатель)

Перемещение фигуры на указанную страницу.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MoveToPage** ( **_Страницы_** **_слева_** **_сверху_**)

 переменная _expression_A, представляющий объект **фигуры** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Page|Обязательное свойство.| **Длинный**|Страница, к которой необходимо переместить фигуру.|
|Слева|Необязательный| **Variant**|Слева от фигуры на странице.|
|Вверх|Необязательный| **Variant**|Верхнюю границу фигуры на странице.|

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **MoveToPage** для перемещения первой фигуры в коллекции **фигур** на первой странице публикации в ту же папку относительный на второй странице публикации.

В этом коде предполагается, что текущей публикации содержит по крайней мере две страницы, и имеется по крайней мере один фигуры на первой странице публикации.




```vb
Public Sub MoveToPage_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 pubShape.MoveToPage 2 
 
End Sub
```


