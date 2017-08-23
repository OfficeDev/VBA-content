---
title: "Свойство WebPageOptions.PublishFileName (издатель)"
keywords: vbapb10.chm544784
f1_keywords: vbapb10.chm544784
ms.prod: publisher
api_name: Publisher.WebPageOptions.PublishFileName
ms.assetid: d3f52a82-8876-303a-2a73-fdb6dd1ff1cb
ms.date: 06/08/2017
ms.openlocfilehash: cb7e488f9217d1730bdc6af0ec0e863a7efeeaf5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionspublishfilename-property-publisher"></a>Свойство WebPageOptions.PublishFileName (издатель)

Возвращает или задает **строку** , представляющую имя файла веб-страницы (в пределах веб-публикации), которое сохраняется как HTML с фильтром.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PublishFileName**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Указание имени файла для веб-страницы является необязательным. При публикации сохраняется в виде HTML с фильтром, Microsoft Publisher автоматически создает имя файла для любой веб-страницы, для которого не указано имя файла. Метод **[SaveAs](document-saveas-method-publisher.md)** объекта **[Document](document-object-publisher.md)** для Сохранение публикации в формате HTML с фильтром.

Файл пользовательских имен используется только в том случае, когда публикации сохраняется в виде HTML с фильтром.

Имена файлов должен быть указан без расширения имени файла.

Включая недопустимые символы (такие как знаки, которые не всегда допустимы в именах файлов, являющихся частью URL-адреса) в поле имя файла возникает ошибка времени выполнения. Недопустимые символы: 


-  символы с кодом выберите значение ниже (десятичное) 48
    
- любой двухбайтовых знаков
    
- следующие символы: \, ?, >, <, |,:, и.
    


Это свойство соответствует текстовое поле **имя файла** в разделе **Опубликовать в Интернете** диалогового окна **Параметры веб-страницы** .


## <a name="example"></a>Пример

В следующем примере задается имя файла и описание второй страницы в активной публикации. Предполагается, что активная публикация — это веб-публикации, содержащий по крайней мере две страницы.


```vb
With ActiveDocument.Pages(2).WebPageOptions 
 .PublishFileName = "CompanyProfile" 
 .Description = "Tailspin Toys Company Profile" 
End With
```


