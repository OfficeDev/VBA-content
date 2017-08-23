---
title: "Свойство Shape.Type (издатель)"
keywords: vbapb10.chm2228307
f1_keywords: vbapb10.chm2228307
ms.prod: publisher
api_name: Publisher.Shape.Type
ms.assetid: bb712dd4-5d81-10e0-9b4c-4af6a09a3c71
ms.date: 06/08/2017
ms.openlocfilehash: 447e9c45d037c7fac56eee4a66898e209cd68bf3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapetype-property-publisher"></a>Свойство Shape.Type (издатель)

Указывает тип фигуры. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Тип**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **типа** может иметь одно из **[PbShapeType](pbshapetype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере форматов выноски тип формы, указанный при выноски фигуры. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub SetCalloutType() 
 With ActiveDocument.Pages(1).Shapes(1) 
 If .Type = pbCallout Then 
 With .Callout 
 .Border = msoTrue 
 .Type = msoCalloutThree 
 End With 
 End If 
 End With 
End Sub
```


