---
title: "Объект стиля текста (издатель)"
keywords: vbapb10.chm6029311
f1_keywords: vbapb10.chm6029311
ms.prod: publisher
api_name: Publisher.TextStyle
ms.assetid: 163ab726-ac44-07d1-ab7b-50061037cc77
ms.date: 06/08/2017
ms.openlocfilehash: 45ff7498833d4625c0fe59e9618e210ad064240b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstyle-object-publisher"></a>Объект стиля текста (издатель)

Представляет один встроенных или пользовательских стилей. Объект **стиля текста** содержит атрибуты стиля (шрифт, стиль шрифта, абзацами и т.д.) в качестве свойства объекта **стиля текста** . Объект **стиля текста** является элементом коллекции **[TextStyles](textstyles-object-publisher.md)** . Коллекция **TextStyles** включает все стили в указанный документ.
 


## <a name="example"></a>Пример

Используйте **TextStyles** (индекс), где индекс — это имя или номер стиля текста, чтобы получить объект **стиля текста** . Должен полностью совпадать правописания и интервал по имени стиля, но не обязательно его правильно.
 

 

 

 
Следующий пример отображает имя стиля и базового стиля первый стиль в коллекции **TextStyles** .
 

 



```
Sub BaseStyleName() 
 With ActiveDocument.TextStyles(1) 
 MsgBox "Style name= " &amp; .Name _ 
 &amp; vbCr &amp; "Base style= " &amp; .BaseStyle 
 End With 
End Sub
```

Используйте метод **[Add](textstyles-add-method-publisher.md)** для создания нового стиля. Применение стиля к диапазону, или несколько абзацев, присвойте свойству **[стиля текста](paragraphformat-textstyle-property-publisher.md)** с именем пользовательских или встроенных стилей. В следующем примере создается новый стиль и применяется к абзац с позиции курсора.
 

 



```
Sub ApplyTextStyle() 
 Dim styNew As TextStyle 
 Dim fntStyle As Font 
 
 'Create a new style 
 Set styNew = ActiveDocument.TextStyles.Add(StyleName:="NewStyle") 
 Set fntStyle = styNew.Font 
 
 'Format the Font object 
 With fntStyle 
 .Name = "Tahoma" 
 .Size = 20 
 .Bold = msoTrue 
 End With 
 
 'Apply the Font object formatting to the new style 
 styNew.Font = fntStyle 
 
 'Apply the new style to the selected paragraph 
 Selection.TextRange.ParagraphFormat.TextStyle = "NewStyle" 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](textstyle-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](textstyle-application-property-publisher.md)|
|[BaseStyle](textstyle-basestyle-property-publisher.md)|
|[Описание](textstyle-description-property-publisher.md)|
|[Шрифт](textstyle-font-property-publisher.md)|
|[Name](textstyle-name-property-publisher.md)|
|[NextParagraphStyle](textstyle-nextparagraphstyle-property-publisher.md)|
|[ParagraphFormat](textstyle-paragraphformat-property-publisher.md)|
|[Родительский раздел](textstyle-parent-property-publisher.md)|

