---
title: "Метод ShapeRange.Distribute (издатель)"
keywords: vbapb10.chm2294017
f1_keywords: vbapb10.chm2294017
ms.prod: publisher
api_name: Publisher.ShapeRange.Distribute
ms.assetid: a145fb46-d7b6-bc3c-b7fd-cdb892fda179
ms.date: 06/08/2017
ms.openlocfilehash: 1201e46407f89ba4b0db8d9285b6f54f2468a2dc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangedistribute-method-publisher"></a>Метод ShapeRange.Distribute (издатель)

Равномерно распределяет фигур в диапазон указанной фигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Распространение** ( **_DistributeCmd_**, **_RelativeTo_**)

 переменная _expression_A, представляющий объект **ShapeRange** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|DistributeCmd|Обязательное свойство.| **MsoDistributeCmd**|Указывает, будет ли фигур распределенных по горизонтали или по вертикали.|
|RelativeTo|Обязательное свойство.| **MsoTriState**|Указывает, следует ли равномерное распределение фигур на все горизонтальный или вертикальный место на странице или в пределах горизонтального или вертикального пространства, который изначально занимает диапазона фигур.|

## <a name="remarks"></a>Заметки

Таким образом, что равно объем пространства между одной фигуры и следующим распределяются фигур. Перекрытием при распределенные по доступное место больших фигур, их распределения, чтобы не равно объем перекрытие между одной фигуры и далее.

Параметр DistributeCmd может иметь одно из следующих **MsoDistributeCmd** константы, описанные в библиотеке типов, Microsoft Office.



| **msoDistributeHorizontally**|| **msoDistributeVertically**| Параметр RelativeTo может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Распределить фигуры в пределах горизонтального или вертикального пространства, который изначально занимает диапазона фигур.|
| **msoTrue**|Равномерное распределение фигур по всей горизонтальное или вертикальное пространство на странице.|
При **msoTrue**RelativeTo фигур распределяются, чтобы расстояние между двумя фигурами внешнего и края страницы совпадает с расстояние между одной фигуры и далее. Если необходимо накладываются фигур, двумя фигурами внешнего перемещаются края страницы.

Когда RelativeTo **msoFalse**два внешних фигур не перемещается; настраиваются только положения внутреннего фигур.

Z порядка фигур влиянию этого метода.


## <a name="example"></a>Пример

В этом примере определяется диапазона фигуры, который содержит все автофигуры на первой странице active публикации и по горизонтали распределяет фигур в этот диапазон.


```vb
' Number of shapes on the page. 
Dim intShapes As Integer 
' Number of AutoShapes on the page. 
Dim intAutoShapes As Integer 
' An array of the names of the AutoShapes. 
Dim arrAutoShapes() As String 
' A looping variable. 
Dim shpLoop As Shape 
' A placeholder variable for the range containing AutoShapes. 
Dim shpRange As ShapeRange 
 
With ActiveDocument.Pages(1).Shapes 
 ' Count all the shapes on the page. 
 intShapes = .Count 
 
 ' Proceed only if there's at least one shape. 
 If intShapes > 1 Then 
 intAutoShapes = 0 
 ReDim arrAutoShapes(1 To intShapes) 
 
 ' Loop through the shapes on the page and add the names 
 ' of any AutoShapes to an array. 
 For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.Type = msoAutoShape Then 
 intAutoShapes = intAutoShapes + 1 
 arrAutoShapes(intAutoShapes) = shpLoop.Name 
 End If 
 Next shpLoop 
 
 ' Proceed only if there's at least one AutoShape. 
 If intAutoShapes > 1 Then 
 ReDim Preserve arrAutoShapes(1 To intAutoShapes) 
 
 ' Create a shape range containing all the AutoShapes. 
 Set shpRange = .Range(Index:=arrAutoShapes) 
 
 ' Distribute the AutoShapes horizontally 
 ' in the space they already occupy. 
 shpRange.Distribute _ 
 DistributeCmd:=msoDistributeHorizontally, RelativeTo:=msoFalse 
 End If 
 End If 
End With 

```


