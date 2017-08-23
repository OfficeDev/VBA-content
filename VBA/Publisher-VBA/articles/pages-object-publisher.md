---
title: "Объект страницы (издатель)"
keywords: vbapb10.chm524287
f1_keywords: vbapb10.chm524287
ms.prod: publisher
api_name: Publisher.Pages
ms.assetid: d6b7262c-015c-dcf3-bff4-0091dd32b78f
ms.date: 06/08/2017
ms.openlocfilehash: 7294f1b256946de97d9c3e16eb7316fd578d89bc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pages-object-publisher"></a>Объект страницы (издатель)

Представляет всех страниц в публикации. Коллекция **Pages** содержит все объекты **[страницы](page-object-publisher.md)** в публикации.
 


## <a name="example"></a>Пример

Используйте метод **[Add](pages-add-method-publisher.md)** для добавления новой страницы публикации. Следующий пример добавляет новую страницу и фигуры active публикации.
 

 

```
Sub AddPageAndShape() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=72, Top:=72, Width:=50, Height:=50) 
 .Fill.ForeColor.RGB = RGB(Red:=128, Green:=50, Blue:=255) 
 .Line.ForeColor.RGB = RGB(Red:=75, Green:=50, Blue:=255) 
 End With 
 End With 
 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](pages-add-method-publisher.md)|
|[AddWizardPage](pages-addwizardpage-method-publisher.md)|
|[FindByPageID](pages-findbypageid-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](pages-application-property-publisher.md)|
|[Count](pages-count-property-publisher.md)|
|[Элемент](pages-item-property-publisher.md)|
|[Родительский раздел](pages-parent-property-publisher.md)|

