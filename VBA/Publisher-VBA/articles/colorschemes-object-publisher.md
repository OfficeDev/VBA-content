---
title: "Объект ColorSchemes (издатель)"
keywords: vbapb10.chm2818047
f1_keywords: vbapb10.chm2818047
ms.prod: publisher
api_name: Publisher.ColorSchemes
ms.assetid: f5002de1-5e91-fc92-eedb-0e13dce57802
ms.date: 06/08/2017
ms.openlocfilehash: 1f300731b1a0d1efdabdc1864282a2dd65960223
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorschemes-object-publisher"></a>Объект ColorSchemes (издатель)

Коллекция всех объектов **[ColorScheme](colorscheme-object-publisher.md)** в Microsoft Publisher. Каждый объект **ColorScheme** представляет цветовая схема, которая представляет собой набор цветов, которые используются в публикации.
 


## <a name="example"></a>Пример

Свойство **[Count](colorschemes-count-property-publisher.md)** возвращает число доступных цветовые схемы Publisher. Следующий пример показывает число цветовые схемы.
 

 

```
Sub CountColorSchemes() 
 MsgBox Application.ColorSchemes.Count 
End Sub
```

Используйте свойство **[Item](colorschemes-item-property-publisher.md)** для возврата определенного цветовая схема из коллекции **ColorSchemes** . ** _Индекса_** аргумент свойства **элемента** может быть номер или имя цветовая схема или константа **PbColorScheme** . В следующем примере задается цветовая схема active публикации для Дикие цветы.
 

 



```
Sub SetColorScheme() 
 ActiveDocument.ColorScheme _ 
 = ColorSchemes.Item(pbColorSchemeWildflower) 
End Sub
```

Используйте свойство **[Name](colorscheme-name-property-publisher.md)** возвращает имя цветовой схемы. В текстовом поле в следующем примере перечисляются все доступные издателю цветовые схемы.
 

 



```
Sub ListColorShemes() 
 
 Dim clrScheme As ColorScheme 
 Dim strSchemes As String 
 
 For Each clrScheme In Application.ColorSchemes 
 strSchemes = strSchemes &amp; clrScheme.Name &amp; vbLf 
 Next 
 ActiveDocument.Pages(1).Shapes.AddTextbox( _ 
 Orientation:=pbTextOrientationHorizontal, _ 
 Left:=72, Top:=72, Width:=400, Height:=500).TextFrame _ 
 .TextRange.Text = strSchemes 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](colorschemes-application-property-publisher.md)|
|[Count](colorschemes-count-property-publisher.md)|
|[Элемент](colorschemes-item-property-publisher.md)|
|[Родительский раздел](colorschemes-parent-property-publisher.md)|

