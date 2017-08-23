---
title: "Объект макетом (издатель)"
keywords: vbapb10.chm655359
f1_keywords: vbapb10.chm655359
ms.prod: publisher
api_name: Publisher.MasterPages
ms.assetid: 3a7e6021-cbe4-4700-018c-c91d2f7d908a
ms.date: 06/08/2017
ms.openlocfilehash: cf4c27dd54af88f19a63b8029fa29634cb135628
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="masterpages-object-publisher"></a>Объект макетом (издатель)

Представляет главной страницы для публикации, после чего будет создан всех страниц в публикации. Объект **макетом** представляет собой коллекцию объектов **[страницы](page-object-publisher.md)** .
 


## <a name="example"></a>Пример

Свойство **[макетом](document-masterpages-property-publisher.md)** используется для возврата объекта **макетом** . В следующем примере добавляется два направляющие на главную страницу, чтобы каждая страница в активной публикации делится на квартала.
 

 

```
Sub ChangeMasterPage() 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 intWidth = .PageSetup.PageWidth 
 intWidth = intWidth / 2 
 intHeight = .PageSetup.PageHeight 
 intHeight = intHeight / 2 
 With .MasterPages(1).RulerGuides 
 .Add Position:=intWidth, _ 
 Type:=pbRulerGuideTypeVertical 
 .Add Position:=intHeight, _ 
 Type:=pbRulerGuideTypeHorizontal 
 End With 
 End With 
End Sub
```

Используйте свойство **[фигур](page-shapes-property-publisher.md)** для работы с автофигуры и текстовых полей на главной странице. В этом примере добавляет небольшой красной фигуре левого верхнего угла главной страницы, которое будет отображаться на каждой странице active публикации.
 

 



```
Sub AddShapeToMasterPage() 
 ActiveDocument.MasterPages(1).Shapes.AddShape(Type:=msoShapeHeart, _ 
 Left:=36, Top:=36, Width:=36, Height:=36).Fill _ 
 .ForeColor.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](masterpages-add-method-publisher.md)|
|[FindByPageID](masterpages-findbypageid-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](masterpages-application-property-publisher.md)|
|[Count](masterpages-count-property-publisher.md)|
|[Элемент](masterpages-item-property-publisher.md)|
|[Родительский раздел](masterpages-parent-property-publisher.md)|

