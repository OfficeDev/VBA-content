---
title: "Метод Shapes.AddLine (издатель)"
keywords: vbapb10.chm2162708
f1_keywords: vbapb10.chm2162708
ms.prod: publisher
api_name: Publisher.Shapes.AddLine
ms.assetid: 43df8878-5640-875f-06e0-37e1feb47b78
ms.date: 06/08/2017
ms.openlocfilehash: 4c41a8f1f1e2711d8f82c3d78532c740f6a968b1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapesaddline-method-publisher"></a>Метод Shapes.AddLine (издатель)

Добавляет новый объект **[фигуры](shape-object-publisher.md)** , представляющее строки для определенной коллекции **[фигур](shapes-object-publisher.md)** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **AddLine** ( **_BeginX_**, **_BeginY_**, **_EndX_**, **_EndY_**)

 переменная _expression_A, представляет собой объект- **фигур** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|BeginX|Обязательное свойство.| **Variant**|X координата начальную точку линии.|
|BeginY|Обязательное свойство.| **Variant**|Начальную точку линии по оси y.|
|EndX|Обязательное свойство.| **Variant**|Координата x конечной точки линии.|
|EndY|Обязательное свойство.| **Variant**|Конечные точки линии по оси y.|

### <a name="return-value"></a>Возвращаемое значение

Shape


## <a name="remarks"></a>Заметки

Для **_BeginX_**, **_BeginY_**, **_EndX_**и **_EndY_** аргументы числовые значения вычисляются в точках; строк может быть в любой устройств, поддерживаемых Microsoft Publisher (например, «2,5 дюйма»).


## <a name="example"></a>Пример

Следующий пример добавляет новую строку для первой страницы active публикации.


```vb
Dim shpLine As Shape 
 
Set shpLine = ActiveDocument.Pages(1).Shapes.AddLine _ 
 (BeginX:=144, BeginY:=144, _ 
 EndX:=180, EndY:=72) 

```


