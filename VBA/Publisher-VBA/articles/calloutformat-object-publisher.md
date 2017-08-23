---
title: "Объект CalloutFormat (издатель)"
keywords: vbapb10.chm2555903
f1_keywords: vbapb10.chm2555903
ms.prod: publisher
api_name: Publisher.CalloutFormat
ms.assetid: 1f54aba3-3872-e668-fe76-1966d1a62cca
ms.date: 06/08/2017
ms.openlocfilehash: ac3cf86f1ea314da8c3c5d6c3fd5f1affd8159c1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="calloutformat-object-publisher"></a>Объект CalloutFormat (издатель)

Содержит свойства и методы, которые применяются к строке выносок.
 


## <a name="example"></a>Пример

Используйте свойство **[выноски](shape-callout-property-publisher.md)** для возврата объекта **CalloutFormat** . В следующем примере добавляется выноски active публикацию добавляет текст в выноске, а затем задает следующие атрибуты для выноски:
 

 

 

 

- Вертикальная черта, отделяющий текст из строки выноски ( **Акцент** свойство)
    
 
- угол между линии выноски и части текстовое поле выноски будет 30 градусов ( **угол** свойство)
    
 
- будет нет границы вокруг текста выноски ( **границы** свойство)
    
 
- линии выноски будет присоединена к верхней части поле выноски (метод **PresetDrop** )
    
 
- линии выноски будут содержаться три сегменты (свойство **Type** )
    
 



```
Sub AddFormatCallout() 
 With ActiveDocument.Pages(1).Shapes.AddCallout(Type:=msoCalloutOne, _ 
 Left:=150, Top:=150, Width:=200, Height:=100) 
 With .TextFrame.TextRange 
 .Text = "This is a callout." 
 With .Font 
 .Name = "Stencil" 
 .Bold = msoTrue 
 .Size = 30 
 End With 
 End With 
 With .Callout 
 .Accent = MsoTrue 
 .Angle = msoCalloutAngle30 
 .Border = MsoFalse 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
 End With 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AutomaticLength](calloutformat-automaticlength-method-publisher.md)|
|[CustomDrop](calloutformat-customdrop-method-publisher.md)|
|[CustomLength](calloutformat-customlength-method-publisher.md)|
|[PresetDrop](calloutformat-presetdrop-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Диакритических знаков](calloutformat-accent-property-publisher.md)|
|[Угол](calloutformat-angle-property-publisher.md)|
|[Приложения](calloutformat-application-property-publisher.md)|
|[AutoAttach](calloutformat-autoattach-property-publisher.md)|
|[AutoLength](calloutformat-autolength-property-publisher.md)|
|[Границы](calloutformat-border-property-publisher.md)|
|[Поместите](calloutformat-drop-property-publisher.md)|
|[DropType](calloutformat-droptype-property-publisher.md)|
|[Разрывов](calloutformat-gap-property-publisher.md)|
|[Длина](calloutformat-length-property-publisher.md)|
|[Родительский раздел](calloutformat-parent-property-publisher.md)|
|[Type](calloutformat-type-property-publisher.md)|

