# -*- coding: utf-8 -*-
"""
Created on Sat Jan  2 12:43:37 2021

Programa Tez dosyasını Word(docx) formatında vermelisiniz!
"""
import docx
from os import chdir
from snowballstemmer import TurkishStemmer
from nltk.tokenize import word_tokenize

# Os kütüphanesini kullanarak, dosyanın bulunduğu dizinin yoluna gidiyoruz.
chdir("C:/Users/Burak/Desktop/tokentry/")


class Tez:
    """
    Sınıfın ismi Tez. Konsolda çalıştırmak için tez'dan nesne türetilmesi gerekmekedir. (x = Tez("TezDosyaAdı.docx"))
    Sınıf türetilirken, dosyanın isminin yazılması 'yol' değişkeninin yapıcı methoda parametre olarak verilmesinden kaynaklıdır.
    """

    def __init__(self, yol):
        self.yol = yol
        self.number = 0
        self.doc = docx.Document(yol)
        self.fullText = []
        self.parsText = []
        self.sourceText = []
        self.source_indis = 0
        self.source_Values = []
        self.searchWord = []
        self.Onsoz = 0
        self.IndısEk = 0
        for i in self.doc.paragraphs:
            self.fullText.append(i.text)

    def FileOrder(self):
        kelime = TurkishStemmer()
        for i in self.fullText:
            if (i == "" or i == "\n"):
                pass
            else:
                self.parsText.append(i)
        for i in self.parsText:
            if (kelime.stemWord(i.lower()) == "kaynak"):
                self.source_indis = self.number
            if (kelime.stemWord(i.lower()) == "önsöz"):
                self.Onsoz = self.number
            if (kelime.stemWord(i.lower()) == "ekler"):
                self.IndısEk = self.number
            else:
                self.number += 1
        print("\t Toplam Boşluk Karakteri Sayısı: ", len(self.fullText) - self.number)
        print("\t Boşluk karakteri olmadan toplam satır sayısı: ", self.number)
        print("\t Kaynakca Başlangıç indisi: ", self.source_indis)
        print("\t Onsoz Başlangıç indisi: ", self.Onsoz)
        print("\t Toplam Yapılan Atıf: ", (self.number - self.source_indis))

    # Düzenlenen Dosyanın tamamını görüntüler.
    def FileOutput(self):
        for i in range(len(self.parsText)):
            print(self.parsText[i])
        return len(self.parsText)

    # Düzenlenen tezi indis araması yapar.
    def FilerowOutput(self, indis):
        try:
            return self.parsText[indis]
        except TypeError:
            print("Görüntülemek istediğiniz satırı giriniz!!")
        except:
            print("Bir şeyler Ters gidiyor")

    # Orjinal tez dosyasında indis ile arama yapar.
    def FileRowOutputOrg(self, indis):
        try:
            return self.fullText[indis]
        except Exception:
            return "Giriş hatası yapıldı. Girişe dikkat edin!"

    # Fonksiyonun parametresine dosyanın satırının tamamı verildiğinde çıktı olarak indis numrasını döndürür
    def RowIndex(self, kelime):
        try:
            return (self.parsText.index(kelime))
        except ValueError:
            return "Kelime Bulunamadı!"

    # Fonksiyonun parametsine aranmak istenen kelime verilir. Çıktı olarak aranan kelime varsa satır numarasını döndürür.
    def WordSearch(self, key):
        try:
            sayi = False
            for i in self.parsText:
                kelime = word_tokenize(str(i))
                for j in kelime:
                    if (j == key):
                        sayi = True
                        self.searchWord.append(self.parsText.index(i))
                    else:
                        pass
            if (sayi == False):
                return "Aranılan Kelime Bulunamadı!"
            else:
                for i in set(self.searchWord):
                    print("Satır Sırası: {}".format(i))
        except Exception:
            return "Giriş İçin Değer Verin!"

    # WordSearch() fonksiyonu çalıştırılıp aranan kelimeler bir listeye atanır bu listeyi WordSearchRows() fonksiyonu çağrılarak görüntülenebilir
    def WordSearchRows(self):
        Keywords = set(self.searchWord)
        for i in Keywords:
            print(self.parsText[i], end="\n \n")

    # Tez dosyası içerisinde ki "Kaynak" bölümünü arayıp kaynakça bölümünde sayfa belirtilmemişse hata satırlarını döndürür!
    def Source_Scanning(self):
        s = 0
        for i in self.parsText[(self.source_indis + 1):]:
            words = word_tokenize(str(i))
            for j in words:
                if (j == "pp" or j == "ss" or j == "sayfa" or j == "page" or j == "syf"):
                    s += 1
                else:
                    pass
            if (s > 0):
                self.source_Values.append(1)
            else:
                self.source_Values.append(0)
            s = 0
        g = len(self.source_Values)
        print(self.source_Values)
        return print("Kaynaklar/Kaynakca bölümünde toplam hata sayısı= ", g)

    # Fonksiyon tez dosyasının giriş Sayfasını kontrol eder.
    def Open_Page(self):
        try:
            dboll = False
            orjinal = ["T.C.", "FIRAT ÜNİVERSİTESİ", "FEN BİLİMLERİ ENSTİTÜSÜ"]
            for i in range(3):
                if (orjinal[i] == self.parsText[i]):
                    dboll = True
                else:
                    dboll = False
            if (dboll == False):
                return "Giriş Etiketinde Uyumsuzluk Tespit Edildi!"
            print("Giriş İşlmeleri Doğru!")
        except Exception:
            pass
        finally:
            pass