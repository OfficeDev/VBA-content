---
title: "Метод CalloutFormat.CustomLength (издатель)"
keywords: vbapb10.chm2490386
f1_keywords: vbapb10.chm2490386
ms.prod: publisher
api_name: Publisher.CalloutFormat.CustomLength
ms.assetid: 855df4af-a02f-fff3-9b12-af886a9788bc
ms.date: 06/08/2017
ms.openlocfilehash: 6c85caee6dbd7e40efe022076a7336830f4c4152
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatcustomlength-method-publisher"></a>Метод CalloutFormat.CustomLength (издатель)

Указывает, что первый сегмент линии выноски (сегмент, подключенного к поле выноски) присваиваться фиксированной длины при каждом перемещении выноске.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CustomLength** ( **_Длина_**)

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Length|Обязательное свойство.| **Variant**|Длина первого сегмента выноски. Числовые значения вычисляются в точках; строк может быть в любой единицы, поддерживаемый Microsoft Publisher (например, «2,5 дюйма»).|

## <a name="remarks"></a>Заметки

Применение этого метода свойству **[AutoLength](calloutformat-autolength-property-publisher.md)** присваивается **значение False** и задает свойство **[Length](calloutformat-length-property-publisher.md)** значение, заданное для аргумента **_длины_** .

Метод **[AutomaticLength](calloutformat-automaticlength-method-publisher.md)** используется для указания, что первый сегмент линии выноски масштабироваться автоматически при каждом перемещении выноске. Применяется только к выноски, чьи строки состоят из нескольких сегментов (типы **msoCalloutThree** и **msoCalloutFour**).


## <a name="example"></a>Пример

В этом примере для переключения между первый сегмент на автоматическое масштабирование и с фиксированной длины строки выноски для первой фигуры в активной публикации. Для обеспечения работы примера этой фигуры должен быть выноске.


```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .AutoLength Then 
 .CustomLength Length:=50 
 Else 
 .AutomaticLength 
 End If 
End With
```


