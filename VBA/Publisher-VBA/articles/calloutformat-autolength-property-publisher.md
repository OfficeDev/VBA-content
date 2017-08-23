---
title: "Свойство CalloutFormat.AutoLength (издатель)"
keywords: vbapb10.chm2490627
f1_keywords: vbapb10.chm2490627
ms.prod: publisher
api_name: Publisher.CalloutFormat.AutoLength
ms.assetid: ed874ec4-d4ce-5e3f-771a-8b3158f40707
ms.date: 06/08/2017
ms.openlocfilehash: 939af381b8aa3d730e648a477e1af45f2e182d81
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformatautolength-property-publisher"></a>Свойство CalloutFormat.AutoLength (издатель)

Возвращает константу **MsoTriState**, указывающее, масштабируется ли первый сегмент линии выноски при перемещении выноске. Применяется только к выноски, чьи строки состоят из нескольких сегментов (типы **msoCalloutThree** и **msoCalloutFour**). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoLength**

 переменная _expression_A, представляет собой объект- **CalloutFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **AutoLength** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Используйте метод [AutomaticLength](calloutformat-automaticlength-method-publisher.md)к этому свойству присвоено значение **msoTrue**, а метод [CustomLength](calloutformat-customlength-method-publisher.md)к этому свойству присвоено значение **msoFalse**.


## <a name="example"></a>Пример

В этом примере для переключения между первый сегмент на автоматическое масштабирование и с фиксированной длины строки выноски для первой фигуры в публикации. Для обеспечения работы примера фигуры должен быть выноске.


```vb
With ActiveDocument.Pages(1).Shapes(1).Callout 
 If .AutoLength Then 
 .CustomLength Length:=50 
 Else 
 .AutomaticLength 
 End If 
End With 

```


