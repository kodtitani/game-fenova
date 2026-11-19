# Fenova – Hızlı Oyun + Günlük Ödüller sürümü

Bu paket Fenova'nın mevcut öğrenci hesapları, arkadaşlık/grup/mesajlaşma, mağaza, lig, futbol bilgi düellosu, sınıfa göre Hızlı Oyun ve Excel etkinlik kaydı yapısını temel alır. Bu sürüm ayrıca günlük hediye ve günlük şans çarkı ekler.

## Günlük Hediye

- Her hesap Türkiye saatiyle (Europe/Istanbul) takvim günü başına bir kez **100 Kua** alabilir.
- Günlük hediye hakkı sunucuda kontrol edilir; tarayıcı yenileyerek veya aynı isteği tekrarlayarak ikinci kez alınamaz.
- Ödül doğrudan mevcut öğrenci cüzdanına eklenir.

## Günlük Şans Çarkı

- Her hesap günde **1 ücretsiz dönüş** hakkına sahiptir.
- Çarkın ödülü istemci tarafından seçilmez; sunucu kriptografik rastgele seçim yapar ve hesabın cüzdanına ekler.
- Ödüller: 50, 100, 150, 250 veya 500 Kua; 1, 2 veya 3 elmas.
- Büyük ödüllerin olasılıkları daha düşüktür. Çark animasyonu, sunucunun seçtiği dilime gider.
- Günlük hediye ile çark hakkı birbirinden bağımsızdır; ikisi de günde birer kez kullanılabilir.

## Veri saklama ve güvenlik

- Günlük hediye/çark hakları `users.json` içindeki her öğrenci kaydının `dailyRewards` alanında tutulur.
- Yeni bir kayıt dosyası oluşturulmaz. Mevcut `users.json` dosyası korunur ve eski hesaplara eksik alanlar otomatik eklenir.
- Excel etkinlik günlüğüne `gunluk_hediye_alindi` ve `gunluk_cark_cevrildi` olayları da gönderilir. Excel dosyasının yeri sabittir: `C:/Users/PC/OneDrive/Desktop/Fenova-Sunucu-Kayit-Sablonu.xlsx`. Sunucu bu dosyaya yazarken Excel kapalı olmalıdır.
- Bir hesabın günlük ödül hakları sunucu tarafından kontrol edilir; istemciden ödül miktarı veya çark dilimi gönderilmez.

## Hızlı Oyun kuralları

- Her sınıfın eşleştirme kuyruğu ayrıdır; farklı sınıflar aynı odaya alınmaz.
- Aynı sınıftan 4 öğrenci bulununca oda kurulur ve 20 saniyelik geri sayım başlar.
- Oda en fazla 8 oyuncu alır; 8 kişi olursa oyun hemen başlar.
- Geri sayım sonunda en az 4 kişi varsa 10 soruluk yarışma başlar.
- Şu anki `questions.json` dosyasında 7. sınıf havuzu bulunur. Diğer sınıflar için aynı formatta ayrı soru havuzları eklenmelidir.

## Windows kurulumu

1. Mevcut Fenova proje klasörünün yedeğini al.
2. ZIP içindeki `server.js` dosyasını proje ana klasörüne, `public/index.html` dosyasını `public` klasörüne kopyala.
3. Mevcut `users.json`, `social.json`, `leagues.json`, `store-codes.json` dosyalarını silme veya boş dosyalarla değiştirme. Mevcut soru bankan daha güncelse `questions.json` üzerine yazma.
4. CMD'de proje klasörüne geçip `npm install` ve ardından `npm start` çalıştır.
5. Tarayıcıda `http://localhost:3000` adresini aç. Yan menüde **Günlük Hediyeler ve Çark** bölümünü göreceksin.

```cmd
cd /d "C:\Users\PC\OneDrive\Fenova-Kahoot-SocketIO-Entegre"
npm install
npm start
```

## Kontrol notu

Sunucu JavaScript'i ve sayfa içindeki istemci JavaScript'i sözdizimi kontrolünden geçirildi. Bu çalışma alanında Express, Socket.IO ve ExcelJS bağımlılıkları kurulu olmadığı için gerçek Windows sunucu/tarayıcı uçtan uca testi yapılamadı. İlk çalıştırmada `npm install` gerekir.
