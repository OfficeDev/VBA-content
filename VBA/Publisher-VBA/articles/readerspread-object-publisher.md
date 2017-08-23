---
title: "Объект ReaderSpread (издатель)"
keywords: vbapb10.chm589823
f1_keywords: vbapb10.chm589823
ms.prod: publisher
api_name: Publisher.ReaderSpread
ms.assetid: 32c55e79-2217-654f-730c-9abaa2cfb9de
ms.date: 06/08/2017
ms.openlocfilehash: 70c0878cbd13978d6128e88150a93ce4498ec1b2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="readerspread-object-publisher"></a>Объект ReaderSpread (издатель)

Представляет ширина чтения (не принтер распространение) для этой страницы. Ширина обычно чтения содержит одну или несколько страниц. Свойства объекта **ReaderSpread** представлены сведения об ли сталкиваются страниц и расположение этих страниц. К примеру в направлена представления страницы, вторая и третья страницы может быть рядом друг с другом или одна под другой.
 


## <a name="example"></a>Пример

Используйте свойство **[ReaderSpread](page-readerspread-property-publisher.md)** для доступа к объекту **ReaderSpread** для страницы. Свойство **[PageCount](readerspread-pagecount-property-publisher.md)** используется для определения, если распространения чтения включает в себя одну или две ориентированные страницы. В этом примере проверяется, если распространения чтения включает в себя меньше, чем две страницы. Если это так, он изменяется ширина для включения две страницы чтения.
 

 

```
Sub SetFacingPages() 
 With ActiveDocument 
 If .Pages.Count >= 2 Then 
 If .Pages(2).ReaderSpread.PageCount < 2 Then _ 
 .ViewTwoPageSpread = True 
 End If 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](readerspread-application-property-publisher.md)|
|[Высота](readerspread-height-property-publisher.md)|
|[Слева](readerspread-left-property-publisher.md)|
|[PageCount](readerspread-pagecount-property-publisher.md)|
|[Страницы](readerspread-pages-property-publisher.md)|
|[Родительский раздел](readerspread-parent-property-publisher.md)|
|[Вверх](readerspread-top-property-publisher.md)|
|[Ширина](readerspread-width-property-publisher.md)|

