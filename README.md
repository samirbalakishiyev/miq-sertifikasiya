# MİQ və Sertifikasiya Hazırlığı

Elvir İsayev · Samir Balakişiyev — WhatsApp 055 795 86 25

Sayt: https://samirbalakishiyev.github.io/miq-sertifikasiya/

## Necə işləyir (My New Vocabulary ilə eyni)
- Tətbiqi ilk dəfə açan hər kəs qeydiyyatdan keçir: ad, soyad, telefon. Qeydiyyat bir dəfədir.
- Qeydiyyatdan 5 gün sonra hesab avtomatik bloklanır və "6 aylıq abunə haqqı 10 AZN" mesajı çıxır.
- Müəllim panelində (S&E → parol) "💳 Ödəniş edib" basılanda blok dərhal açılır və 6 ay bloklanmır.
- Müəllimlər heç vaxt bloklanmır.

## Google Cədvəl (bir dəfəlik quraşdırma)
1. Google Drive → Yeni → Google Cədvəl. Adı: "MİQ və Sertifikasiya".
2. Uzantılar → Apps Script. Oradakı kodu silin, `Kod.gs` faylının hamısını yapışdırın.
3. Yuxarıdakı `BURAYA-PANEL-SIFRESI` sözünün yerinə panel parolunu yazın. 💾 Yadda saxlayın.
4. Yerləşdir (Deploy) → Yeni yerləşdirmə → Növ: Veb tətbiq.
   - İcra edən: Mən
   - Giriş: Hər kəs
5. "Yerləşdir" → icazələri verin → çıxan Veb tətbiq URL-ini (…/exec) kopyalayın.
6. Bu URL `index.html`-də `SHEET_URL` kimi yazılır.

Kod.gs dəyişəndə: Yerləşdir → Yerləşdirmələri idarə et → ✏️ → Versiya: Yeni versiya → Yerləşdir (URL dəyişmir).
