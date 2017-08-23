---
title: "Свойство ShapeRange.HasTextFrame (издатель)"
keywords: vbapb10.chm2293858
f1_keywords: vbapb10.chm2293858
ms.prod: publisher
api_name: Publisher.ShapeRange.HasTextFrame
ms.assetid: 8a3b4f3b-3282-686b-f4fe-abf2d7677b3e
ms.date: 06/08/2017
ms.openlocfilehash: f40e542b349807843fe45d2737eed1ea35c82653
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangehastextframe-property-publisher"></a>Свойство ShapeRange.HasTextFrame (издатель)

Указывает, имеет ли указанный фигуры **TextFrame** объекта, связанного с ним. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasTextFrame**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Если свойство **HasTextFrame** имеет значение true, клиенты должны проверять значение свойства **HasText** объекта **TextFrame** , чтобы определить, если в форме любого текста.

Значение свойства **HasTextFrame** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Указанный фигура не имеет **TextFrame** объекта, связанного с ним.|
| **msoTriStateMixed**| Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**| Указанный фигура имеет **TextFrame** объекта, связанного с ним.|

## <a name="example"></a>Пример

В этом примере проверяется всех фигур в выделение, и если нет рамок текста, связанные с ними, они являются по левому краю.


```vb
Sub MoveLeft() 
 
 Dim shpAll As ShapeRange 
 
 Set shpAll = Application.ActiveDocument.Selection.ShapeRange 
 If shpAll.HasTextFrame = msoFalse Then 
 shpAll.Align msoAlignLefts, msoTrue 
 End If 
 
End Sub
```


