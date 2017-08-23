---
title: "Объект LayoutGuides (издатель)"
keywords: vbapb10.chm1179647
f1_keywords: vbapb10.chm1179647
ms.prod: publisher
api_name: Publisher.LayoutGuides
ms.assetid: 7430c1c4-c7f5-d9b6-cea8-b21fe9e2905f
ms.date: 06/08/2017
ms.openlocfilehash: 6048616621e46af5928b10c8d146ac8f36e96049
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="layoutguides-object-publisher"></a>Объект LayoutGuides (издатель)

Представляет таблицы измерения, которое отображается на страницах публикации наложения помочь при размещении элементы дизайна.
 


## <a name="example"></a>Пример

Используйте свойство **[LayoutGuides](document-layoutguides-property-publisher.md)** объекта **Document** для получения объекта **LayoutGuides** . Используйте объект **LayoutGuide** свойств поля и свойства **строк** и **столбцов** для количества строк и столбцов отображаются в направляющие разметки и где они отображаются на странице.
 

 

 

 
В этом примере устанавливаются поля активной презентации для двух дюйма.
 

 



```
With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
End With
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](layoutguides-application-property-publisher.md)|
|[ColumnGutterWidth](layoutguides-columngutterwidth-property-publisher.md)|
|[Столбцы](layoutguides-columns-property-publisher.md)|
|[GutterCenterlines](layoutguides-guttercenterlines-property-publisher.md)|
|[HorizontalBaseLineOffset](layoutguides-horizontalbaselineoffset-property-publisher.md)|
|[HorizontalBaseLineSpacing](layoutguides-horizontalbaselinespacing-property-publisher.md)|
|[MarginBottom](layoutguides-marginbottom-property-publisher.md)|
|[MarginLeft](layoutguides-marginleft-property-publisher.md)|
|[MarginRight](layoutguides-marginright-property-publisher.md)|
|[MarginTop](layoutguides-margintop-property-publisher.md)|
|[MirrorGuides](layoutguides-mirrorguides-property-publisher.md)|
|[Родительский раздел](layoutguides-parent-property-publisher.md)|
|[RowGutterWidth](layoutguides-rowgutterwidth-property-publisher.md)|
|[Строк](layoutguides-rows-property-publisher.md)|
|[VerticalBaseLineOffset](layoutguides-verticalbaselineoffset-property-publisher.md)|
|[VerticalBaseLineSpacing](layoutguides-verticalbaselinespacing-property-publisher.md)|

