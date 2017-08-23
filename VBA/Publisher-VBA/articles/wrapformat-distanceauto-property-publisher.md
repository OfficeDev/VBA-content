---
title: "Свойство WrapFormat.DistanceAuto (издатель)"
keywords: vbapb10.chm786437
f1_keywords: vbapb10.chm786437
ms.prod: publisher
api_name: Publisher.WrapFormat.DistanceAuto
ms.assetid: 8b4e6b93-6e68-5c4a-2164-1a88ca0a633e
ms.date: 06/08/2017
ms.openlocfilehash: 369e5e5b7241caeeb838b32c0eb56050e49b1698
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="wrapformatdistanceauto-property-publisher"></a>Свойство WrapFormat.DistanceAuto (издатель)

Возвращает или задает константой **MsoTriState** , указывающее, соответствующем расстоянии между встроенная фигура и окружающим текстом автоматически рассчитывается ли. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DistanceAuto**

 переменная _expression_A, представляет собой объект- **WrapFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **DistanceAuto** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Привязка фигуры в зависимости от поля текстовое поле, которое он перекрывается не настраиваются.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTriStateToggle**| Задайте значение, которое выполняется переключение между **msoTrue** и **msoFalse**значение свойства.|
| **msoTrue**|По умолчанию. Привязка фигуры автоматически настраивается в зависимости от поля текстовое поле, которое он перекрывается. |

## <a name="example"></a>Пример

В следующем примере задается фигуры на странице активная публикация, чтобы ее края не настраивается автоматически на основании расстояние от окружающим текстом.


```vb
Sub SetDistanceAutoProperty() 
 With ActiveDocument.Pages(1).Shapes(1).TextWrap 
 .Type = pbWrapTypeSquare 
 .DistanceAuto = msoFalse 
 End With 
End Sub
```


