# ExcelBuilder
ExcelBuilder, C# ile Excel dosyalarını oluşturmayı, düzenlemeyi ve veri aktarımını kolaylaştıran bir yardımcı sınıftır. Bu sınıf, EPPlus kütüphanesini kullanarak Excel dosyalarıyla etkileşim sağlar.

## Özellikler
* Excel dosyası oluşturma veya mevcut bir şablona dayalı olarak düzenleme
* Satır ve sütun ekleme, kopyalama ve silme işlemleri
* Hücre verilerini ayarlama, biçimlendirme ve formül eklemeyi destekleme
* Veri koleksiyonlarını hızlı bir şekilde Excel tablosuna aktarma
* Excel dosyasını farklı formatlarda (xlsx, csv) kaydetme
* Excel dosyasından veri aktarımı ve dönüştürme

## Kullanım
```cs
// Yeni bir ExcelBuilder örneği oluşturma
var builder = new ExcelBuilder("Sheet1");

// Başlık ve veri ekleme
builder.SetData("A1", "My Title")
    .SetData("A2", "Data 1")
    .SetData("B2", "Data 2");

// Verileri koleksiyondan aktarma
List<MyDataModel> data = GetDataFromSomewhere();
builder.SetDataList("A4", data);

// Excel dosyasını kaydetme
builder.Build("C:\\Documents", "MyExcelFile");

// Excel dosyasından veri almak
var importedData = ExcelBuilder.ImportToEntity<MyDataModel>("C:\\Documents\\MyExcelFile.xlsx");
```

Bu örnekler, ExcelBuilder sınıfının temel kullanımını göstermektedir. Sınıf, daha birçok işlevselliği destekler ve farklı senaryolara uyarlanabilir.

## Notlar
* ExcelBuilder sınıfı, EPPlus kütüphanesine dayanmaktadır. Bu nedenle, projenize EPPlus kütüphanesini eklemeniz gerekmektedir.
* Sınıfın içinde yer alan metodlar ve parametreler, gerekli uyarlama ve genişletmeler yapılarak ihtiyaçlara göre özelleştirilebilir.


Bu ExcelBuilder sınıfı, C# ile Excel işlemleri gerçekleştirmek isteyen geliştiriciler için bir temel oluşturmayı hedeflemektedir. İhtiyaçlarınıza göre bu sınıfı genişletebilir veya uyarlayabilirsiniz.
