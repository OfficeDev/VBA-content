---
title: "Метод ShapeRange.Align (издатель)"
keywords: vbapb10.chm2294016
f1_keywords: vbapb10.chm2294016
ms.prod: publisher
api_name: Publisher.ShapeRange.Align
ms.assetid: ef522d47-3fc7-cfca-5b9a-44ff020f8b31
ms.date: 06/08/2017
ms.openlocfilehash: 23fd8068805fa12d078c0fceb8625705d3dc9ae4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangealign-method-publisher"></a>Метод ShapeRange.Align (издатель)

Выравнивает всех фигур на указанный объект **ShapeRange** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выравнивание** ( **_AlignCmd_**, **_RelativeTo_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|AlignCmd|Обязательное свойство.| **MsoAlignCmd**|Задает способ выравнивания фигур.|
|RelativeTo|Обязательное свойство.| **MsoTriState**|Указывает, выполняется ли выравнивание фигур, относящиеся к странице или друг с другом.|

## <a name="remarks"></a>Заметки

Если параметр RelativeTo является **msoFalse** и диапазона фигуры содержит только одну фигуру, возникает ошибка.

Параметр AlignCmd может иметь одно из **MsoAlignCmd** константы в библиотеке типов, Microsoft Office.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoAlignBottoms**|Выравнивает фигур по нижнему краю. Если _RelativeTo_ **msoFalse**, нижних фигур определяет строку, с которым выровненные другие фигуры.|
| **msoAlignCenters**|Выравнивает фигур на вертикальной линии через их центры. Если _RelativeTo_ **msoFalse**, в строке промежуточное слева - и правые фигуры выравнивания фигур.|
| **msoAlignLefts**|Выравнивает фигур по левому краю. Если _RelativeTo_ **msoFalse**, самые левые фигуры определяет строку, с которым выровненные другие фигуры.|
| **msoAlignMiddles**|Выравнивает фигур на горизонтальную линию через их центры. Если _RelativeTo_ **msoFalse**, в строке Промежуточное начало - и нижних фигур выравнивания фигур.|
| **msoAlignRights**| **msoAlignRights** Выравнивает фигур по правому краю. Если _RelativeTo_ **msoFalse**, правые фигуры определяет строку, с которым выровненные другие фигуры.|
| **msoAlignTops**| Выравнивает фигур по верхнему краю. Если _RelativeTo_ **msoFalse**, верхняя фигура определяет строку, с которым выровненные другие фигуры.|
Параметр RelativeTo может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Выравнивает фигур относительно друг с другом.|
| **msoTrue**|Выравнивает фигур, относящиеся к странице.|

## <a name="example"></a>Пример

Следующий пример выравнивание всех фигур на первой странице active публикации на вертикальной линии через их центры.


```vb
ActiveDocument.Pages(1).Shapes.Range.Align _ 
 AlignCmd:=msoAlignCenters, _ 
 RelativeTo:=msoTrue 

```


