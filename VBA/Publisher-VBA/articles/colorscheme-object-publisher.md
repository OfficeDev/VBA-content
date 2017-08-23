---
title: "Объект ColorScheme (издатель)"
keywords: vbapb10.chm2752511
f1_keywords: vbapb10.chm2752511
ms.prod: publisher
api_name: Publisher.ColorScheme
ms.assetid: b4e554ef-f043-c963-e175-b7d5ba95c636
ms.date: 06/08/2017
ms.openlocfilehash: e7701a1630b2efe6b95dcb8830a6beae4773bb6c
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorscheme-object-publisher"></a>Объект ColorScheme (издатель)

Представляет цветовая схема, которая представляет собой набор из восьми цветов, используемые для различных элементов публикации. Каждый цвет соответствует **[ColorFormat](colorformat-object-publisher.md)** объект. Объект **ColorScheme** является элементом коллекции **[ColorSchemes](colorschemes-object-publisher.md)** . Коллекция **ColorSchemes** содержит все доступные для Microsoft Publisher цветовые схемы.
 


## <a name="example"></a>Пример

Свойство **[ColorScheme](document-colorscheme-property-publisher.md)** объекта **[Document](document-object-publisher.md)** для возврата цветовая схема для текущей публикации. В следующем примере задается значение заполнения трех фигур на первой странице возвращаемое значение (в формате RGB) из трех **ColorScheme** цветов.
 

 

```
Sub ReturnColorsAndApplyToShapes() 
 Dim lngAccent1 As Long 
 Dim lngAccent2 As Long 
 Dim lngAccent3 As Long 
 
 With ActiveDocument 
 With .ColorScheme 
 lngAccent1 = .Colors(pbSchemeColorAccent1).RGB 
 lngAccent2 = .Colors(pbSchemeColorAccent2).RGB 
 lngAccent3 = .Colors(pbSchemeColorAccent3).RGB 
 End With 
 With .Pages(1) 
 .Shapes(1).Fill.ForeColor.RGB = lngAccent1 
 .Shapes(2).Fill.ForeColor.RGB = lngAccent2 
 .Shapes(3).Fill.ForeColor.RGB = lngAccent3 
 End With 
 End With 
 
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
|[Приложения](colorscheme-application-property-publisher.md)|
|[Цвета](colorscheme-colors-property-publisher.md)|
|[Name](colorscheme-name-property-publisher.md)|
|[Родительский раздел](colorscheme-parent-property-publisher.md)|

