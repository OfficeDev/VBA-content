---
title: "Свойство WebCommandButton.DataFileFormat (издатель)"
keywords: vbapb10.chm3932169
f1_keywords: vbapb10.chm3932169
ms.prod: publisher
api_name: Publisher.WebCommandButton.DataFileFormat
ms.assetid: 7594b575-b39f-3cd4-d0b9-c13c04299345
ms.date: 06/08/2017
ms.openlocfilehash: cbaf6876e67c68c3cb91ec9e734607a8a76c42ad
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcommandbuttondatafileformat-property-publisher"></a>Свойство WebCommandButton.DataFileFormat (издатель)

Задает или возвращает константу **PbSubmitDataFormatType** , который представляет формат для использования при сохранении в файл данных веб-форм. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DataFileFormat**

 переменная _expression_A, представляет собой объект- **WebCommandButton** .


### <a name="return-value"></a>Возвращаемое значение

PbSubmitDataFormatType


## <a name="remarks"></a>Заметки

Значение свойства **DataFileFormat** может иметь одно из **[PbSubmitDataFormatType](pbsubmitdataformattype-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В этом примере задается Microsoft Publisher процесс данных веб-форм путем сохранения файла с разделителями-запятыми на одном веб-сервере, как форма будет сохранена. (Обратите внимание на то, имя файла, заменены допустимое имя файла для работы этого примера).


```vb
Sub WebDataFile() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalSaveOnServer 
 .DataFileFormat = pbSubmitDataFormatCSV 
 .DataFileName = "Filename" 
 End With 
End Sub
```


