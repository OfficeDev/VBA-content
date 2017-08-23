---
title: "Метод CalloutFormat.AutomaticLength (издатель)"
keywords: vbapb10.chm2490384
f1_keywords: vbapb10.chm2490384
ms.prod: publisher
api_name: Publisher.CalloutFormat.AutomaticLength
ms.assetid: 3772ad87-9808-5f25-0b9c-cdd7b1392ca1
ms.date: 06/08/2017
ms.openlocfilehash: 963a878fd5df6667a992307d039f92edb399388a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatautomaticlength-method-publisher"></a>Метод CalloutFormat.AutomaticLength (издатель)

Указывает, что первый сегмент линии выноски (сегмент, подключенного к поле выноски) масштабироваться автоматически, при перемещении выноске.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutomaticLength**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


## <a name="remarks"></a>Заметки

При вызове этого метода присваивает свойству **[AutoLength](calloutformat-autolength-property-publisher.md)** указанного объекта значение **msoTrue**.

Метод **[CustomLength](calloutformat-customlength-method-publisher.md)** используется для указания, что первый сегмент линии выноски сохранять фиксированной длины, возвращаемой свойством **[Длина](calloutformat-length-property-publisher.md)** при каждом перемещении выноске. Применяется только к выноски, чьи строки состоят из нескольких сегментов (типы **msoCalloutThree** и **msoCalloutFour**).


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


