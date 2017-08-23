---
title: "Объект Font (издатель)"
keywords: vbapb10.chm5439487
f1_keywords: vbapb10.chm5439487
ms.prod: publisher
api_name: Publisher.Font
ms.assetid: 992fda94-2820-d665-0d78-efd4b5434731
ms.date: 06/08/2017
ms.openlocfilehash: faeb751ff4b393ff53970ab11f19d7bdb0c26f64
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="font-object-publisher"></a>Объект Font (издатель)

Содержит атрибуты шрифта (шрифт, размер шрифта, цвета и т. п.) для объекта.


## <a name="example"></a>Пример

Свойство **[шрифта](http://msdn.microsoft.com/library/80d7177a-fef9-c3fd-f559-94644a2ba0f7%28Office.15%29.aspx)** возвращает объект **Font** . Следующие инструкции применяет жирное форматирование для выделения.


```
Sub BoldText() 
 Selection.TextRange.Font.Bold = True 
End Sub
```

В следующем примере форматируются в первый абзац active публикации как Arial и курсив 24 точки.




```
Sub FormatText() 
 Dim txtRange As TextRange 
 Set txtRange = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
 With txtRange.Font 
 .Bold = True 
 .Name = "Arial" 
 .Size = 24 
 End With 
End Sub
```

В следующем примере изменяется форматирования стиля Заголовок 2 в активной публикации на Arial и полужирным шрифтом.




```
Sub FormatStyle() 
 With ActiveDocument.TextStyles("Normal").Font 
 .Name = "Tahoma" 
 .Italic = True 
 .Size = 15 
 End With 
End Sub
```

Также вы можете продублировать объект **Font** , используя свойство **[повторяющихся](http://msdn.microsoft.com/library/545dbfdb-4cd5-99b1-1ba3-b723e8d7b827%28Office.15%29.aspx)** . В следующем примере создается новый стиль знака с форматирование символов из выделения в дополнение к курсивного начертания. Форматирование выделенного фрагмента не изменяется.




```
Sub DuplicateFont() 
 Dim fntNew As Font 
 Set fntNew = Selection.TextRange.Font.Duplicate 
 fntNew.Italic = True 
 ActiveDocument.TextStyles.Add(StyleName:="Italics").Font = fntNew 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Дублирующиеся](http://msdn.microsoft.com/library/26ae64bc-036e-5c19-cbac-99f11da7fb60%28Office.15%29.aspx)|
|[GetScriptName](http://msdn.microsoft.com/library/332860de-33fa-7d6a-ac42-28c39856cff7%28Office.15%29.aspx)|
|[Увеличьте размеры](http://msdn.microsoft.com/library/41d48db2-4a0d-6efc-80c5-c6f035e9e6ff%28Office.15%29.aspx)|
|[Сброс](http://msdn.microsoft.com/library/7a81d7f9-4db9-3ce1-188d-2b4719b57fff%28Office.15%29.aspx)|
|[SetScriptName](http://msdn.microsoft.com/library/f1f2c01e-098c-1afd-0e64-1d563c1ca626%28Office.15%29.aspx)|
|[Сжатие](http://msdn.microsoft.com/library/c5626ef2-5351-ab49-bf86-690587daed1f%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[AllCaps](http://msdn.microsoft.com/library/e8394f91-de31-0075-51ac-8a372023f0ce%28Office.15%29.aspx)|
|[Приложения](http://msdn.microsoft.com/library/e4721e0f-c591-3ac6-319d-2e753f1b375a%28Office.15%29.aspx)|
|[AttachedToText](http://msdn.microsoft.com/library/23b0519a-9f35-fa25-752a-4942e8161edd%28Office.15%29.aspx)|
|[AutomaticPairKerningThreshold](http://msdn.microsoft.com/library/f5f43a19-7227-b25d-9322-84a79596c525%28Office.15%29.aspx)|
|[Полужирный шрифт](http://msdn.microsoft.com/library/3b9ba2b0-c319-9d08-9a36-5b292046962e%28Office.15%29.aspx)|
|[BoldBi](http://msdn.microsoft.com/library/f3a9fa27-6c9c-4d77-0f0d-962afa211d9d%28Office.15%29.aspx)|
|[ContextualAlternates](http://msdn.microsoft.com/library/4737d43a-4ab8-0ae7-ce45-7be62f4aae6e%28Office.15%29.aspx)|
|[DiacriticColor](http://msdn.microsoft.com/library/6e9c816e-c7ae-c559-6b35-150a5abb820c%28Office.15%29.aspx)|
|[ExpandUsingKashida](http://msdn.microsoft.com/library/ecf3a170-5f07-379e-ff56-504beb770308%28Office.15%29.aspx)|
|[Заполните поля](http://msdn.microsoft.com/library/c38ac8a3-2673-c968-9fcb-ebd5545d4da4%28Office.15%29.aspx)|
|[Свечения](http://msdn.microsoft.com/library/72fb3acb-e405-a03a-1e12-88b775551f7f%28Office.15%29.aspx)|
|[Курсив](http://msdn.microsoft.com/library/c55c0bfa-a365-86ac-4cfb-f6911dadd0af%28Office.15%29.aspx)|
|[ItalicBi](http://msdn.microsoft.com/library/604e776c-92b0-6e5b-2599-ab879c61a78a%28Office.15%29.aspx)|
|[Кернинг](http://msdn.microsoft.com/library/756fe3fa-9bf3-be16-2dd1-5b8fb0ec6496%28Office.15%29.aspx)|
|[В то же время](http://msdn.microsoft.com/library/17847824-8761-42b7-8d0c-00345e8b5de8%28Office.15%29.aspx)|
|[Строка](http://msdn.microsoft.com/library/56add50f-85f4-0c65-cc64-3a68000d9428%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/03561991-5456-aee3-4c04-56a2520a4d6e%28Office.15%29.aspx)|
|[NumberStyle](http://msdn.microsoft.com/library/e4adedac-e3a5-4a85-8825-ba24c32dca60%28Office.15%29.aspx)|
|[Родительский раздел](http://msdn.microsoft.com/library/c02da1ef-014f-3c83-a2a8-8afa474be4e1%28Office.15%29.aspx)|
|[Position](http://msdn.microsoft.com/library/24573faf-1627-3b10-5a8e-2f76a9f8831d%28Office.15%29.aspx)|
|[Отражения](http://msdn.microsoft.com/library/e426d097-4839-6949-147c-f84b230bdfb7%28Office.15%29.aspx)|
|[Масштабирование](http://msdn.microsoft.com/library/4ff0c484-12f8-38e3-72fd-dfd34507aec1%28Office.15%29.aspx)|
|[Размер](http://msdn.microsoft.com/library/485f68fe-c6d7-8288-042e-fc4c35c37b2d%28Office.15%29.aspx)|
|[SizeBi](http://msdn.microsoft.com/library/1e9100e7-efa4-a7aa-69af-39c550a0b046%28Office.15%29.aspx)|
|[SmallCaps](http://msdn.microsoft.com/library/ab50b850-f371-7d8e-0c19-00ad68e700f0%28Office.15%29.aspx)|
|[Зачеркивание](http://msdn.microsoft.com/library/fa4bca2d-b43d-4d2b-901f-858e277df520%28Office.15%29.aspx)|
|[StylisticAlternates](http://msdn.microsoft.com/library/cfb46152-4a54-27df-0a77-1e8b7fd3a711%28Office.15%29.aspx)|
|[StylisticSets](http://msdn.microsoft.com/library/0d25fbf3-8d68-c10f-0d1b-526314700329%28Office.15%29.aspx)|
|[Подстрочный знак](http://msdn.microsoft.com/library/9992fdcc-dd60-b2f7-307b-99b10dc7debb%28Office.15%29.aspx)|
|[Надстрочный знак](http://msdn.microsoft.com/library/582c02c9-4dcb-f826-8ec0-e9e10702f717%28Office.15%29.aspx)|
|[Swash](http://msdn.microsoft.com/library/71537393-167a-f9e3-e3b3-ae743fdbb0ff%28Office.15%29.aspx)|
|[TextShadow](http://msdn.microsoft.com/library/052948b2-205b-6934-d659-17e3b17f8590%28Office.15%29.aspx)|
|[ThreeD](http://msdn.microsoft.com/library/947691ab-5b38-8b3c-3615-a205a27ba4c3%28Office.15%29.aspx)|
|[Отслеживание](http://msdn.microsoft.com/library/c703a5ec-e8d7-36ce-ac50-d41265ce92db%28Office.15%29.aspx)|
|[TrackingPreset](http://msdn.microsoft.com/library/818e6efd-a1b3-1ccd-1dc1-29c0a8ded7f2%28Office.15%29.aspx)|
|[Подчеркивание](http://msdn.microsoft.com/library/a01a943e-274d-725e-3f78-aa76c51d5c46%28Office.15%29.aspx)|
|[UseDiacriticColor](http://msdn.microsoft.com/library/368d3599-b0b0-1790-0ce0-13f1936bccb0%28Office.15%29.aspx)|

