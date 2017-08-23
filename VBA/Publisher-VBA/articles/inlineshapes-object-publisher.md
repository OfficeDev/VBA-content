---
title: "Объект InlineShapes (издатель)"
keywords: vbapb10.chm5832703
f1_keywords: vbapb10.chm5832703
ms.prod: publisher
api_name: Publisher.InlineShapes
ms.assetid: 1a6d1e8f-0be0-102e-af6c-a1cee53eae02
ms.date: 06/08/2017
ms.openlocfilehash: a6c4219099616a94c6a45c36e5103e27c2aab5c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="inlineshapes-object-publisher"></a>Объект InlineShapes (издатель)

Содержит коллекцию объектов **[фигуры](shape-object-publisher.md)** , которые представляют объекты в графических, где **Shape.IsInline** имеет **значение True**. Коллекции фигур ограничен фигур в диапазоне заданный текст.
 


## <a name="remarks"></a>Заметки

Набор **InlineShapes** доступен только на объекте **TextRange** . С помощью **TextFrame.Story.TextRange.InlineShapes** возвращает всех встроенных фигур в рамке, включая те, которые находятся в переполнения. С помощью **TextFrame.TextRange.InlineShapes** возвращает только видимые встроенных фигур в фрагмент текста, а не указанные в переполнения.
 

 
Коллекции **InlineShapes** также можно получить доступ из **Document.Stories ( _i_ ). TextRange**, где i — индекс на активную страницу публикации.
 

 
Коллекция **InlineShapes** недоступна в коллекции **Page.Shapes** , включая его автономные **ShapeRange**.
 

 

## <a name="example"></a>Пример

Свойство **[InlineShapes](textrange-inlineshapes-property-publisher.md)** объекта **[TextRange](textrange-object-publisher.md)** используется для возврата коллекции **InlineShapes** . В следующем примере выполняется поиск первой фигуры в текстовом поле на странице публикации и добавляет текст в конец диапазона текст в текстовом поле при наличии более одного встроенного фигуры в диапазон текста.
 

 

```
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.TextRange 
 If .InlineShapes.Count > 1 Then 
 .InsertAfter (" There is more than one inline shape in this text box.") 
 End If 
End With
```

Используйте свойство **InlineShapes** (индекс) для возвращения одного встроенного фигуры. В следующем примере производится поиск третий встроенная фигура в текстовое поле и зеркальное отражение по вертикали.
 

 



```
Dim theShape As Shape 
 
Set theShape = ActiveDocument.Pages(1).Shapes(1) 
 
With theShape.TextFrame.Story.TextRange 
 With .InlineShapes(3) 
 .Flip (msoFlipVertical) 
 End With 
End With
```

Используйте метод **[диапазона](shapes-range-method-publisher.md)** возвращает объект **[ShapeRange](shaperange-object-publisher.md)** , содержащий все элементы из коллекции **InlineShapes** . Массив индексов строк или отдельный индекс или строка может передается как параметр свойства **диапазон** для выбора конкретного фигур или фигуры в диапазоне. В следующем примере задается переменная **ShapeRange** равно коллекцию встроенных фигур, существующих в текстовом поле. Каждая фигура встроенного в диапазоне изменяется каким-либо образом. В этом примере предполагает первую фигуру на странице текстовое поле, которое содержит три встроенных фигур.
 

 



```
Dim theRange As ShapeRange 
 
Set theRange = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.Story.TextRange.InlineShapes.Range 
 
With theRange 
 .Item(1).Flip msoFlipVertical 
 .Item(2).MoveOutOfTextFlow 
 .Item(3).Delete 
End With
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](inlineshapes-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](inlineshapes-application-property-publisher.md)|
|[Count](inlineshapes-count-property-publisher.md)|
|[Родительский раздел](inlineshapes-parent-property-publisher.md)|
|[Range](inlineshapes-range-property-publisher.md)|

