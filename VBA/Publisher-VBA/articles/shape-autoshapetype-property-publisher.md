---
title: "Свойство Shape.AutoShapeType (издатель)"
keywords: vbapb10.chm2228274
f1_keywords: vbapb10.chm2228274
ms.prod: publisher
api_name: Publisher.Shape.AutoShapeType
ms.assetid: f469dc31-a620-5561-ce57-fbff8a5536c0
ms.date: 06/08/2017
ms.openlocfilehash: 12594a2d2852491fd56344bfecf430e968f42433
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapeautoshapetype-property-publisher"></a>Свойство Shape.AutoShapeType (издатель)

Возвращает или задает константой **MsoAutoShapeType**, определяющее тип объекта **Shape** автофигуры.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AutoShapeType**

 переменная _expression_A, представляющий объект **фигуры** .


## <a name="remarks"></a>Заметки

Значение свойства **AutoShapeType** может иметь одно из ** [MsoAutoShapeType](http://msdn.microsoft.com/library/7e6fe414-2b25-56d7-a678-b6e718329118%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Автофигуры соответствуют **фигуры,** несмотря на то, что свойство **AutoShapeType** для фигур издателя, не являющегося также вернет значение. WordArt OLE, веб-форму элемент управления, в таблице объектов и изображение кадра должен возвращать **msoShapeMixed** как их значение свойства **AutoShapeType** . Текстовые рамки должен возвращать **msoShapeRectangle** как свойство их **AutoShapeType** .


## <a name="example"></a>Пример

В этом примере преобразует выбранный объект **автофигуры** на молнию, если это основа и 5-конечная звезда если он не установлен. В данном примере для правильного выполнения необходимо иметь **автофигуры** объекта, выбранного в активной публикации.


```vb
Sub ShapeShift() 
 
 Dim srShift As ShapeRange 
 
 Set srShift = Application.ActiveDocument.Selection.ShapeRange 
 If srShift.AutoShapeType = msoShapeHeart Then 
 srShift.AutoShapeType = msoShapeLightningBolt 
 Else 
 srShift.AutoShapeType = msoShape5pointStar 
 End If 
 
End Sub
```


