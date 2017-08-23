---
title: "Объект View (издатель)"
keywords: vbapb10.chm393215
f1_keywords: vbapb10.chm393215
ms.prod: publisher
api_name: Publisher.View
ms.assetid: a865cf48-cd68-5789-6855-c09c05b7634b
ms.date: 06/08/2017
ms.openlocfilehash: 4d061716fb5b48707378322a5d8cbf289651006f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="view-object-publisher"></a>Объект View (издатель)

Содержит атрибуты, представление (Показать все, затенение полей, сетку и т. п.) окна или панели.
 


## <a name="example"></a>Пример

Свойство **[ActiveView](document-activeview-property-publisher.md)** возвращает объект **View** . В следующем примере задается параметр масштаба.
 

 

```
Sub ZoomFitSelection() 
 ActiveDocument.ActiveView.Zoom = pbZoomFitSelection 
End Sub
```

В приведенных ниже примерах и, соответственно, масштаба активного представления.
 

 



```
Sub ViewZoomIn() 
 ActiveDocument.ActiveView.ZoomIn 
End Sub 
 
Sub ViewZoomOut() 
 ActiveDocument.ActiveView.ZoomOut 
End Sub
```

Следующий пример active прокрутку до указанного фигуры.
 

 



```
Sub ScrollToShape() 
 Dim shpOne As Shape 
 
 Set shpOne = ActiveDocument.Pages(1).Shapes(1) 
 ActiveDocument.ActiveView.ScrollShapeIntoView Shape:=shpOne 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ScrollShapeIntoView](view-scrollshapeintoview-method-publisher.md)|
|[ZoomIn](view-zoomin-method-publisher.md)|
|[ZoomOut](view-zoomout-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[ActivePage](view-activepage-property-publisher.md)|
|[Приложения](view-application-property-publisher.md)|
|[Родительский раздел](view-parent-property-publisher.md)|
|[Показать](view-zoom-property-publisher.md)|

