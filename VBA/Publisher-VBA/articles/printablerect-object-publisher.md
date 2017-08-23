---
title: "Объект PrintableRect (издатель)"
keywords: vbapb10.chm7602175
f1_keywords: vbapb10.chm7602175
ms.prod: publisher
api_name: Publisher.PrintableRect
ms.assetid: fd99e9d4-81d9-63ae-78ca-f7a16b031239
ms.date: 06/08/2017
ms.openlocfilehash: 4780bb421eba59bdcdeab762cdd2f2454ff62cbd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="printablerect-object-publisher"></a>Объект PrintableRect (издатель)

Представляет область листа, в течение которого указанного печать. Область печати определяется принтера на основе указанного размера листа. Подготовленных к печати прямоугольника листа принтера не следует путать с область внутри поля страницы публикации; может быть меньше или больше страницы публикации.
 


## <a name="remarks"></a>Заметки

В тех случаях, когда идентичны sheet принтера и размер страницы публикации страница публикации располагается на листе принтера и ни один из метки печати, даже в том случае, если они выбраны.
 

 

## <a name="example"></a>Пример

Используйте свойство **[PrintableRect](printer-printablerect-property-publisher.md)** объекта **[AdvancedPrintOptions](advancedprintoptions-object-publisher.md)** возвращает объект **PrintableRect** . В следующем примере возвращается ограничениях подготовленных к печати прямоугольника в листе принтера active публикации.
 

 

```
Sub ListPrintableRectBoundaries() 
 
With ActiveDocument.AdvancedPrintOptions.PrintableRect 
 
 Debug.Print "Printable area is " &amp; _ 
 PointsToInches(.Width) &amp; _ 
 " by " &amp; PointsToInches(.Height) &amp; " inches." 
 Debug.Print "Left Boundary: " &amp; PointsToInches(.Left) &amp; _ 
 " inches (from left)." 
 Debug.Print "Right Boundary: " &amp; PointsToInches(.Left + .Width) &amp; _ 
 " inches (from left)." 
 Debug.Print "Top Boundary: " &amp; PointsToInches(.Top) &amp; _ 
 " inches(from top)." 
 Debug.Print "Bottom Boundary: " &amp; PointsToInches(.Top + .Height) &amp; _ 
 " inches(from top)." 
 
End With 
 
End Sub 

```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](printablerect-application-property-publisher.md)|
|[Высота](printablerect-height-property-publisher.md)|
|[Слева](printablerect-left-property-publisher.md)|
|[Родительский раздел](printablerect-parent-property-publisher.md)|
|[Вверх](printablerect-top-property-publisher.md)|
|[Ширина](printablerect-width-property-publisher.md)|

