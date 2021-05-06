using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Sinav_Planlama_Cetveli;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Sınav_Planlama_Cetveli
{
    public partial class Form1 : Form
    {
        string dosyaYolu; //Düzenleyeceğimiz dosyanın yolu
        string connectionString; // Connection string
        DateTime sinavBaslangic; // Sınavın başlangıç saati
        DateTime sinavAralik; //Adayların bir sınavı ile diğer sınavı arasındaki aralık
        int planlamaCetveliSatirSayisi = 0;
        int sinavSuresi;

        //Orman Üretim Sınavı Listeler
        List<Aday> lstOrmanUretimA1B1B2T1Liste = new List<Aday>();
        List<Aday> lstOrmanUretimB1P1Liste = new List<Aday>();
        List<Aday> lstOrmanUretimB2P1Liste = new List<Aday>(); 
        //Orman Üretim Sınavı Listeler
        List<Aday> lstOrmanYetistirmeA1T1Liste = new List<Aday>();
        List<Aday> lstOrmanYetistirmeB1P1Liste = new List<Aday>();
        List<Aday> lstOrmanYetistirmeB2P1Liste = new List<Aday>();
        List<Aday> lstOrmanYetistirmeB3P1Liste = new List<Aday>();

        //Çıktı vereceğimiz json dosyası için listeler
        List<Aday> tmpOrmanUretimA1T1full = new List<Aday>();
        List<Aday> tmpOrmanUretimB1T1full = new List<Aday>();
        List<Aday> tmpOrmanUretimB2T1full = new List<Aday>();
        List<Aday> tmpOrmanUretimB1P1full = new List<Aday>();
        List<Aday> tmpOrmanUretimB2P1full = new List<Aday>();

        List<Aday> tmpOrmanYetistirmeA1T1full = new List<Aday>();
        List<Aday> tmpOrmanYetistirmeB1P1full = new List<Aday>();
        List<Aday> tmpOrmanYetistirmeB2P1full = new List<Aday>();
        List<Aday> tmpOrmanYetistirmeB3P1full = new List<Aday>();



        string sinavYeri; //Sınava girilecek tarih ve yer
        public Form1()
        {
            InitializeComponent();
        }


        /// <summary>
        /// Bu fonksiyon seçilen dosyanın yolunu döndürmeye yarar.
        /// </summary>
        string DosyaAdiGetir()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                return ofd.FileName;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Bu fonksiyon bir SQL sorgusunu dataGridView1'e yansıtmaya yarar.
        /// </summary>
        public void ExcelSorgu(string sqlSorgusu)
        {
            try
            {
                if (dosyaYolu == null)
                {
                    MessageBox.Show("Lütfen Geçerli Bir Yol Giriniz");
                }
                
                else
                {
                    OleDbConnection baglanti = new OleDbConnection(connectionString);
                    baglanti.Open();
                    DataTable dt = new DataTable();
                    OleDbDataAdapter da = new OleDbDataAdapter(sqlSorgusu, baglanti);
                    da.Fill(dt);
                    ormanUretimData.DataSource = dt.DefaultView;
                    baglanti.Close();
                    tabControl1.Visible = true;
                    progressBar1.Visible = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /// <summary>
        /// Bu fonksiyon bir SQL sorgusundan çıkan tek bir sütunu bir string listesi halinde döndürmeye yarar.
        /// </summary>
        List<string> ListFromDatabase(string sorgu)
        {
            List<string> items = new List<string>();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                var cmd = new OleDbCommand(sorgu, connection);
                cmd.CommandType = CommandType.Text;
                connection.Open();
                using (OleDbDataReader objReader = cmd.ExecuteReader())
                {
                    if (objReader.HasRows)
                    {
                        while (objReader.Read())
                        {
                            string item = objReader.GetValue(0).ToString();
                            items.Add(item);
                        }
                    }
                }
            }

            return items;
        }

        /// <summary>
        /// Bu fonksiyon bir SQL sorgusunun çıktısı olan tek bir değeri string halinde döndürmeye yarar.
        /// </summary>
        public string StringFromDatabase(string sorgu)
        {
            string deger = "";
            OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
            baglanti.Open();
            using (OleDbCommand command = new OleDbCommand(sorgu, baglanti))
            {
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    deger = reader.GetString(0);
                }

                if (deger == null)
                {

                }
            }
            baglanti.Close();
            return deger;
        }


        /// <summary>
        /// Bu fonksiyon girilen tarih listesindeki veriyi, GG:AA:YYYY şeklinde değilse listeden çıkarmaya yarar. Ayrıca sıralamayı tarihe göre değiştirir
        /// </summary>
        List<string> TarihFormat(List<string> tarihler)
        {
            List<string> temp = new List<string>();
            List<DateTime> gecici = new List<DateTime>();
            foreach(var tarih in tarihler)
            {
                try
                {
                    gecici.Add(DateTime.ParseExact(tarih, "d.M.yyyy", System.Globalization.CultureInfo.InvariantCulture));
                }

                catch (FormatException) { }
            }
            gecici = gecici.OrderBy(x => x.Date).ToList();
            foreach(var item in gecici)
            {
                temp.Add(item.ToString("d.M.yyyy", System.Globalization.CultureInfo.InvariantCulture));
            }
            return temp;
        }
        
        /// <summary>
         /// Bu fonksiyon adayların sınava atanmasını sağlar.
         /// </summary>
        List<Aday> SinavFromDatabase(string sorgu)
        {
            List<Aday> sinav = new List<Aday>();

            using (OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'"))
            {
                baglanti.Open();

                using (OleDbCommand command = new OleDbCommand(sorgu, baglanti))
                {
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Aday temp = new Aday(sinavSuresi)
                            {
                                tcKimlikNo = reader.GetString(0),
                                adiSoyadi = reader.GetString(1),
                                degerlendirici = reader.GetString(2),
                            };
                            temp.degerlendirici = StringFromDatabase("SELECT A.degerlendiriciadisoyadi FROM [Three$] A  WHERE A.degerlendiricitc = '" + temp.degerlendirici + "'").ToUpper();
                            sinav.Add(temp);
                        }
                    }
                }
            }
            return sinav;
        }
        List<Aday> SinavFromDatabase(string sorgu, bool tarihliMi)
        {
            List<Aday> sinav = new List<Aday>();

            using (OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'"))
            {
                baglanti.Open();

                using (OleDbCommand command = new OleDbCommand(sorgu, baglanti))
                {
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Aday temp = new Aday(sinavSuresi)
                            {
                                tcKimlikNo = reader.GetString(0),
                                adiSoyadi = reader.GetString(1),
                                degerlendirici = reader.GetString(2),
                                girisTarihi = reader.GetValue(3).ToString()
                            };
                            temp.degerlendirici = StringFromDatabase("SELECT A.degerlendiriciadisoyadi FROM [Three$] A  WHERE A.degerlendiricitc = '" + temp.degerlendirici + "'").ToUpper();
                            sinav.Add(temp);
                        }
                    }
                }
            }
            return sinav;
        }
        /// <summary>
        /// Bu fonksiyon ihtiyaç duyulan durumlarda alanların bir string yardımı ile çekilmesini sağlar.
        /// </summary>
        string GetFieldValue(Aday src, string fieldName)
        {
            return src.GetType().GetField(fieldName).GetValue(src).ToString();
        }

        /// <summary>
        /// Bu fonksiyon tekrar eden adayların listeden silinmesini sağlar.
        /// </summary>
        List<Aday> DeleteDuplicates(string sorgu)
        {
            var temp = new List<Aday>();
            lstOrmanUretimA1B1B2T1Liste = SinavFromDatabase(sorgu);
            Aday gecici = new Aday(sinavSuresi);
            temp.Add(gecici);
            bool tekrar;
            for (int i = 0; i < lstOrmanUretimA1B1B2T1Liste.Count; i++)
            {
                tekrar = false;
                for (int j = 0; j < temp.Count; j++)
                {
                    if (lstOrmanUretimA1B1B2T1Liste[i].tcKimlikNo == temp[j].tcKimlikNo)
                        tekrar = true;
                }

                if (!tekrar)
                {
                    temp.Add(lstOrmanUretimA1B1B2T1Liste[i]);
                }
            }
            temp.RemoveAt(0);
            return temp;
        }
        bool ilkSinavMi;

        /// <summary>
        /// Bu fonksiyon teorik sınavının sınav sürelerinin ayarlamasının yapılmasını ve teorik sınavlarının birine veya ikisine girecek kişinin hangi sınavlara gireceğini gösterecek işaretler eklenmesini sağlar.
        /// </summary>
        void TeorikSaatveIsimAyarlama(string tarih)
        {
            List<Aday> teorikA1 = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%A1' AND sinavtarihi = '" + tarih + "'");
            List<Aday> teorikB1 = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%B1' AND sinavtarihi = '" + tarih + "'");
            List<Aday> teorikB2 = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%B2' AND sinavtarihi = '" + tarih + "'");
            DateTime saat = DateTime.Parse(txtOrmanUretimBaslangicSaati.Text);
            for (int i = 0; i < lstOrmanUretimA1B1B2T1Liste.Count; i++)
            {
                string sinavlar = "";
                ilkSinavMi = true;
                int sure = 10;
                for (int j = 0; j < teorikA1.Count; j++)
                {
                    if (lstOrmanUretimA1B1B2T1Liste[i].tcKimlikNo == teorikA1[j].tcKimlikNo)
                    {
                        teorikA1[j].girisSaati = saat;
                        sure += 10;
                        saat = saat.AddMinutes(sure);
                        ilkSinavMi = false;
                        sinavlar += "A1/";
                        break;
                    }
                }
                for (int j = 0; j < teorikB1.Count; j++)
                {
                    if (lstOrmanUretimA1B1B2T1Liste[i].tcKimlikNo == teorikB1[j].tcKimlikNo)
                    {
                        sure += 5;
                        if (ilkSinavMi)
                        {
                            teorikB1[j].girisSaati = saat;
                            saat = saat.AddMinutes(sure);
                            ilkSinavMi = false;
                        }
                        else
                        {
                            teorikB1[j].girisSaati = saat;
                            saat = saat.AddMinutes(5);
                        }
                        sinavlar += "B1/";
                        break;

                    }
                }
                for (int j = 0; j < teorikB2.Count; j++)
                {
                    if (lstOrmanUretimA1B1B2T1Liste[i].tcKimlikNo == teorikB2[j].tcKimlikNo)
                    {
                        sure += 5;
                        if (ilkSinavMi)
                        {
                            teorikB2[j].girisSaati = saat;
                            saat = saat.AddMinutes(sure);
                            ilkSinavMi = false;
                        }
                        else
                        {
                            teorikB2[j].girisSaati = saat;
                            saat = saat.AddMinutes(5);
                        }
                        sinavlar += "B2";
                        break;
                    }
                }
                if(lstOrmanUretimA1B1B2T1Liste[i].tcKimlikNo == "" || sinavlar == "A1/B1/B2")
                {
                    lstOrmanUretimA1B1B2T1Liste[i].sinavSuresi = int.Parse(txtSinavSuresi.Text);
                }
                else
                {
                    lstOrmanUretimA1B1B2T1Liste[i].sinavSuresi = sure;
                }

                switch (sinavlar)
                {
                    case "A1/B1/B2":
                        break;
                    case "A1/":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (A1T1)";
                        break;
                    case "B1/":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (B1T1)";
                        break;
                    case "B2":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (B2T1)";
                        break;
                    case "A1/B1/":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (A1T1/B1T1)";
                        break;
                    case "A1/B2":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (A1T1/B2T1)";
                        break;
                    case "B1/B2":
                        lstOrmanUretimA1B1B2T1Liste[i].adiSoyadi += " (B1T1/B2T1)";
                        break;
                    default:
                        break;
                }
            }

            foreach (var item in teorikA1)
                tmpOrmanUretimA1T1full.Add(item);
            foreach (var item in teorikB1)
                tmpOrmanUretimB1T1full.Add(item);
            foreach (var item in teorikB2)
                tmpOrmanUretimB2T1full.Add(item);
        }

        /// <summary>
        /// Bu fonksiyon giriş değeri olarak verilen listenin istenilen birim kadar kaydırılmasını sağlar
        /// </summary>
        List<Aday> Rotate(List<Aday> list, int offset)
        {
            return list.Skip(offset).Concat(list.Take(offset)).ToList();
        }

        /// <summary>
        /// Bu fonksiyon adayların sınava giriş saatlerinin atanmasını sağlar.
        /// </summary>
        void GirisSaatiAyarlama(List<Aday> list)
        {
            DateTime saat = sinavBaslangic;
            foreach(var item in list)
            {
                item.girisSaati = saat;
                saat = saat.AddMinutes(item.sinavSuresi);
            }
        }
        /// <summary>
        /// Bu fonksiyon saatlerin çakışmaması için gereken gruplamanın yapılmasını sağlar
        /// </summary>
        void OrmanUretimGruplayici()
        {
            List<Aday> kesmeTemp = new List<Aday>();
            List<Aday> surutmeTemp = new List<Aday>();
            List<Aday> eslesemeyenler = new List<Aday>();
            List<Aday> geciciKopya;

            

            //Kesme ve Sürütme sınavlarının listelerinin geçici olarak tutulacağı boş listeler tanımlanıyor.
            for (int i = 0; i < lstOrmanUretimB1P1Liste.Count; i++)
            {
                Aday temp = new Aday(sinavSuresi);
                kesmeTemp.Add(temp);
            }
            for (int i = 0; i < lstOrmanUretimB2P1Liste.Count; i++)
            {
                Aday temp = new Aday(sinavSuresi);
                surutmeTemp.Add(temp);
            }

            //Bu boş listelerin saatleri ayarlanıyor böylelikle saat karşılaştırması yapabiliriz.
            GirisSaatiAyarlama(kesmeTemp);
            GirisSaatiAyarlama(surutmeTemp);


            //Kesme Sınavı listesinin saatlerinin ayarlanması
            for (int i = 0; i < lstOrmanUretimB1P1Liste.Count; i++)
            {
                bool varMi = false;
                for (int j = 0; j < lstOrmanUretimA1B1B2T1Liste.Count; j++)
                {
                    int sayac = i;
                    //Eğer kişi hem teorik sınavına hem de performans sınavına giriyorsa:
                    if(lstOrmanUretimB1P1Liste[i].tcKimlikNo == lstOrmanUretimA1B1B2T1Liste[j].tcKimlikNo)
                    {
                        varMi = true;
                        geciciKopya = KopyasiniAl(lstOrmanUretimB1P1Liste);

                        //Saatlerin arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanUretimA1B1B2T1Liste[j].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            //Liste elemanını bir alt satıra çekiyoruz, böylelikle saatini de değiştirmiş oluyoruz.
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac+1)%geciciKopya.Count;
                        }
                        
                        //Eğer girmek istediğimiz indexteki tc kimlik numarası boş ise:
                        if (kesmeTemp[sayac].tcKimlikNo == "")
                        {
                            kesmeTemp[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve indexteki tc kimlik numarası boş olana kadar dön
                           
                            while ((kesmeTemp[sayacSayaci].girisSaati - lstOrmanUretimA1B1B2T1Liste[j].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || kesmeTemp[sayacSayaci].tcKimlikNo != "")
                            {
                                var deneme = (kesmeTemp[sayacSayaci].girisSaati - lstOrmanUretimA1B1B2T1Liste[j].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || kesmeTemp[sayacSayaci].tcKimlikNo != "";
                                sayacSayaci = (sayacSayaci+1) % kesmeTemp.Count;
                            }

                            kesmeTemp[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }
                }
                //Eğer hiçbir listede yok ise bu kişileri istediğimiz saate atabiliriz. Bu "joker" kişileri eşleşemeyenler listesine atıyoruz.
                if(!varMi)
                {
                    eslesemeyenler.Add(lstOrmanUretimB1P1Liste[i]);
                }

            }

            int degisken = 0;
            //Eşleşemeyenler listesinde kimse kalmayıncaya kadar dön
            while (eslesemeyenler.Any())
            {
                //Tüm listeyi dönüyoruz. Kimseyle eşleşmeyen dolayısı ile istediğimiz saate alabildiğimiz kişileri bu boş yerlere atıyoruz.
                if(kesmeTemp[degisken].tcKimlikNo == "")
                {
                    kesmeTemp[degisken].SaatsizDegerleriAl(eslesemeyenler[0]);
                    eslesemeyenler.RemoveAt(0);
                }
                degisken = (degisken + 1) % kesmeTemp.Count;
            }
            //Oluşan geçici listeyi kaydediyoruz
            lstOrmanUretimB1P1Liste = kesmeTemp;

            //Sıra sürütme sınavının listesine geldi

            for (int i = 0; i < lstOrmanUretimB2P1Liste.Count; i++)
            {
                int sayac = i;
                bool kesmedeVar = false;
                int kesmeIndex = 0;
                bool teorikteVar = false;
                int teorikIndex = 0;

                //Eğer kişi kesme sınavının listesinde var ise, indexini alıyoruz.
                for (int j = 0; j < lstOrmanUretimB1P1Liste.Count; j++)
                {
                    if (lstOrmanUretimB2P1Liste[i].tcKimlikNo == lstOrmanUretimB1P1Liste[j].tcKimlikNo)
                    {
                        kesmedeVar = true;
                        kesmeIndex = j;
                        break;
                    }
                }
                //Eğer kişi teorik sınavının listesinde var ise, indexini alıyoruz.
                for (int k = 0; k < lstOrmanUretimA1B1B2T1Liste.Count; k++)
                {
                    if (lstOrmanUretimB2P1Liste[i].tcKimlikNo == lstOrmanUretimA1B1B2T1Liste[k].tcKimlikNo)
                    {
                        teorikteVar = true;
                        teorikIndex = k;
                        break;
                    }
                }

                //Eğer hem kesmede hem teorikte var ise 
                if (kesmedeVar && teorikteVar)
                {
                    geciciKopya = KopyasiniAl(lstOrmanUretimB2P1Liste);

                    //Her iki sınavın arasında 2 saat oluncaya kadar dön
                    while ((geciciKopya[sayac].girisSaati - lstOrmanUretimA1B1B2T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanUretimB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                    {
                        Swap(geciciKopya, sayac, sayac + 1);
                        sayac = (sayac + 1) % geciciKopya.Count;
                    }

                    //Eğer girmeye çalıştığımız index boş ise:
                    if (surutmeTemp[sayac].tcKimlikNo == "")
                    {
                        surutmeTemp[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }

                    //Değil ise
                    else
                    {
                        int sayacSayaci = sayac;
                        //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                        while (((surutmeTemp[sayacSayaci].girisSaati - lstOrmanUretimA1B1B2T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (surutmeTemp[sayacSayaci].girisSaati - lstOrmanUretimB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString())) || surutmeTemp[sayacSayaci].tcKimlikNo != "")
                        {
                            sayacSayaci = (sayacSayaci + 1) % surutmeTemp.Count;
                        }

                        surutmeTemp[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }

                    geciciKopya.Clear();


                }
                //Kişi sadece kesme sınavının listesinde var ise
                else if (kesmedeVar)
                {
                    geciciKopya = KopyasiniAl(lstOrmanUretimB2P1Liste);
                    //Sınavların arasında 2 saat oluncaya kadar dön
                    while ((geciciKopya[sayac].girisSaati - lstOrmanUretimB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                    {
                        Swap(geciciKopya, sayac, sayac + 1);
                        sayac = (sayac + 1) % geciciKopya.Count;
                    }

                    //Eğer girmeye çalıştığımız index boş ise:
                    if (surutmeTemp[sayac].tcKimlikNo == "")
                    {
                        surutmeTemp[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }

                    //Değil ise
                    else
                    {
                        int sayacSayaci = sayac;
                        //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                        while ((surutmeTemp[sayacSayaci].girisSaati - lstOrmanUretimB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || surutmeTemp[sayacSayaci].tcKimlikNo != "")
                        {
                            sayacSayaci = (sayacSayaci + 1) % surutmeTemp.Count;
                        }

                        surutmeTemp[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }

                    geciciKopya.Clear();


                }

                //Eğer sadece teorik sınavının listesinde var ise
                else if (teorikteVar)
                {
                    geciciKopya = KopyasiniAl(lstOrmanUretimB2P1Liste);
                    //Sınavların arasında 2 saat oluncaya kadar dön
                    while ((geciciKopya[sayac].girisSaati - lstOrmanUretimA1B1B2T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                    {
                        Swap(geciciKopya, sayac, sayac + 1);
                        sayac = (sayac + 1) % geciciKopya.Count;
                    }
                    //Eğer girmeye çalıştığımız index boş ise:
                    if (surutmeTemp[sayac].tcKimlikNo == "")
                    {
                        surutmeTemp[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }
                    //Değil ise
                    else
                    {
                        int sayacSayaci = sayac;
                        //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                        while ((surutmeTemp[sayacSayaci].girisSaati - lstOrmanUretimA1B1B2T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || surutmeTemp[sayacSayaci].tcKimlikNo != "")
                        {
                            sayacSayaci = (sayacSayaci + 1) % surutmeTemp.Count;
                        }

                        surutmeTemp[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                    }

                    geciciKopya.Clear();
                }
                //Eğer diğer listelerde adı yok ise eşleşemeyenler listesine ekleniyor.
                else
                {
                    eslesemeyenler.Add(lstOrmanUretimB2P1Liste[i]);
                }

            }

            degisken = 0;
            while (eslesemeyenler.Any())
            {
                //Indexteki tc kimlik numarası eğer boş ise, bu kişileri bu yerlere ekliyoruz
                if (surutmeTemp[degisken].tcKimlikNo == "")
                {
                    surutmeTemp[degisken].SaatsizDegerleriAl(eslesemeyenler[0]);
                    eslesemeyenler.RemoveAt(0);
                }
                degisken = (degisken + 1) % surutmeTemp.Count;
            }
            //Oluşan geçici listeyi kaydediyoruz
            lstOrmanUretimB2P1Liste = surutmeTemp;
        }

        /// <summary>
        /// Bu fonksiyon saatlerin çakışmaması için gereken gruplamanın yapılmasını sağlar
        /// </summary>
        void OrmanYetistirmeGruplayici()
        {
            List<Aday> tempB1P1 = new List<Aday>();
            List<Aday> tempB2P1 = new List<Aday>();
            List<Aday> tempB3P1 = new List<Aday>();
            List<Aday> eslesemeyenler = new List<Aday>();
            List<Aday> geciciKopya;



            //Kesme ve Sürütme sınavlarının listelerinin geçici olarak tutulacağı boş listeler tanımlanıyor.
            for (int i = 0; i < lstOrmanYetistirmeB1P1Liste.Count; i++)
            {
                Aday temp = new Aday(sinavSuresi);
                tempB1P1.Add(temp);
            }
            for (int i = 0; i < lstOrmanYetistirmeB2P1Liste.Count; i++)
            {
                Aday temp = new Aday(sinavSuresi);
                tempB2P1.Add(temp);
            }
            for (int i = 0; i < lstOrmanYetistirmeB3P1Liste.Count; i++)
            {
                Aday temp = new Aday(sinavSuresi);
                tempB3P1.Add(temp);
            }

            //Bu boş listelerin saatleri ayarlanıyor böylelikle saat karşılaştırması yapabiliriz.
            GirisSaatiAyarlama(tempB1P1); GirisSaatiAyarlama(tempB2P1); GirisSaatiAyarlama(tempB3P1);

            if (lstOrmanYetistirmeB1P1Liste.Any())
            {
                for (int i = 0; i < lstOrmanYetistirmeB1P1Liste.Count; i++)
                {
                    bool varMi = false;
                    for (int j = 0; j < lstOrmanYetistirmeA1T1Liste.Count; j++)
                    {
                        int sayac = i;
                        //Eğer kişi hem teorik sınavına hem de performans sınavına giriyorsa:
                        if (lstOrmanYetistirmeB1P1Liste[i].tcKimlikNo == lstOrmanYetistirmeA1T1Liste[j].tcKimlikNo)
                        {
                            varMi = true;
                            geciciKopya = KopyasiniAl(lstOrmanYetistirmeB1P1Liste);

                            //Saatlerin arasında 2 saat oluncaya kadar dön
                            while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[j].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                            {
                                //Liste elemanını bir alt satıra çekiyoruz, böylelikle saatini de değiştirmiş oluyoruz.
                                Swap(geciciKopya, sayac, sayac + 1);
                                sayac = (sayac + 1) % geciciKopya.Count;
                            }

                            //Eğer girmek istediğimiz indexteki tc kimlik numarası boş ise:
                            if (tempB1P1[sayac].tcKimlikNo == "")
                            {
                                tempB1P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                            }
                            //Değil ise
                            else
                            {
                                int sayacSayaci = sayac;
                                //Saatler uygun ve indexteki tc kimlik numarası boş olana kadar dön
                                while ((tempB1P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[j].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB1P1[sayacSayaci].tcKimlikNo != "")
                                {
                                    sayacSayaci = (sayacSayaci + 1) % tempB1P1.Count;
                                }

                                tempB1P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                            }

                            geciciKopya.Clear();
                        }
                    }
                    //Eğer hiçbir listede yok ise bu kişileri istediğimiz saate atabiliriz. Bu "joker" kişileri eşleşemeyenler listesine atıyoruz.
                    if (!varMi)
                    {
                        eslesemeyenler.Add(lstOrmanYetistirmeB1P1Liste[i]);
                    }

                }

                int degisken = 0;
                //Eşleşemeyenler listesinde kimse kalmayıncaya kadar dön
                while (eslesemeyenler.Any())
                {
                    //Tüm listeyi dönüyoruz. Kimseyle eşleşmeyen dolayısı ile istediğimiz saate alabildiğimiz kişileri bu boş yerlere atıyoruz.
                    if (tempB1P1[degisken].tcKimlikNo == "")
                    {
                        tempB1P1[degisken].SaatsizDegerleriAl(eslesemeyenler[0]);
                        eslesemeyenler.RemoveAt(0);
                    }
                    degisken = (degisken + 1) % tempB1P1.Count;
                }
                //Oluşan geçici listeyi kaydediyoruz
                lstOrmanYetistirmeB1P1Liste = tempB1P1;
            }


            //Sıra sürütme sınavının listesine geldi

            if (lstOrmanYetistirmeB2P1Liste.Any())
            {
                for (int i = 0; i < lstOrmanYetistirmeB2P1Liste.Count; i++)
                {
                    int sayac = i;
                    bool kesmedeVar = false;
                    int kesmeIndex = 0;
                    bool teorikteVar = false;
                    int teorikIndex = 0;

                    //Eğer kişi kesme sınavının listesinde var ise, indexini alıyoruz.
                    for (int j = 0; j < lstOrmanYetistirmeB1P1Liste.Count; j++)
                    {
                        if (lstOrmanYetistirmeB2P1Liste[i].tcKimlikNo == lstOrmanYetistirmeB1P1Liste[j].tcKimlikNo)
                        {
                            kesmedeVar = true;
                            kesmeIndex = j;
                            break;
                        }
                    }
                    //Eğer kişi teorik sınavının listesinde var ise, indexini alıyoruz.
                    for (int k = 0; k < lstOrmanYetistirmeA1T1Liste.Count; k++)
                    {
                        if (lstOrmanYetistirmeB2P1Liste[i].tcKimlikNo == lstOrmanYetistirmeA1T1Liste[k].tcKimlikNo)
                        {
                            teorikteVar = true;
                            teorikIndex = k;
                            break;
                        }
                    }

                    //Eğer hem kesmede hem teorikte var ise 
                    if (kesmedeVar && teorikteVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB2P1Liste);

                        //Her iki sınavın arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB2P1[sayac].tcKimlikNo == "")
                        {
                            tempB2P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while (((tempB2P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB2P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString())) || tempB2P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB2P1.Count;
                            }

                            tempB2P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();


                    }
                    //Kişi sadece kesme sınavının listesinde var ise
                    else if (kesmedeVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB2P1Liste);
                        //Sınavların arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB2P1[sayac].tcKimlikNo == "")
                        {
                            tempB2P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB2P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[kesmeIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB2P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB2P1.Count;
                            }

                            tempB2P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();


                    }

                    //Eğer sadece teorik sınavının listesinde var ise
                    else if (teorikteVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB2P1Liste);
                        //Sınavların arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }
                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB2P1[sayac].tcKimlikNo == "")
                        {
                            tempB2P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB2P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB2P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB2P1.Count;
                            }

                            tempB2P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }
                    //Eğer diğer listelerde adı yok ise eşleşemeyenler listesine ekleniyor.
                    else
                    {
                        eslesemeyenler.Add(lstOrmanYetistirmeB2P1Liste[i]);
                    }

                }

                int degisken = 0;
                while (eslesemeyenler.Any())
                {
                    //Indexteki tc kimlik numarası eğer boş ise, bu kişileri bu yerlere ekliyoruz
                    if (tempB2P1[degisken].tcKimlikNo == "")
                    {
                        tempB2P1[degisken].SaatsizDegerleriAl(eslesemeyenler[0]);
                        eslesemeyenler.RemoveAt(0);
                    }
                    degisken = (degisken + 1) % tempB2P1.Count;
                }
                //Oluşan geçici listeyi kaydediyoruz
                lstOrmanYetistirmeB2P1Liste = tempB2P1;

            }



            if (lstOrmanYetistirmeB3P1Liste.Any())
            {
                for (int i = 0; i < lstOrmanYetistirmeB3P1Liste.Count; i++)
                {
                    int sayac = i;
                    bool teorikteVar = false;
                    bool b1p1deVar = false;
                    bool b2p1deVar = false;
                    int teorikIndex = 0;
                    int b1p1Index = 0;
                    int b2p1Index = 0;

                    //Eğer kişi teorik sınavının listesinde var ise, indexini alıyoruz.
                    for (int j = 0; j < lstOrmanYetistirmeA1T1Liste.Count; j++)
                    {
                        if (lstOrmanYetistirmeB3P1Liste[i].tcKimlikNo == lstOrmanYetistirmeA1T1Liste[j].tcKimlikNo)
                        {
                            teorikteVar = true;
                            teorikIndex = j;
                            break;
                        }
                    }

                    //Eğer kişi kesme sınavının listesinde var ise, indexini alıyoruz.
                    for (int j = 0; j < lstOrmanYetistirmeB1P1Liste.Count; j++)
                    {
                        if (lstOrmanYetistirmeB3P1Liste[i].tcKimlikNo == lstOrmanYetistirmeB1P1Liste[j].tcKimlikNo)
                        {
                            b1p1deVar = true;
                            b1p1Index = j;
                            break;
                        }
                    }

                    for(int j = 0; j < lstOrmanYetistirmeB2P1Liste.Count; j++)
                    {
                        if (lstOrmanYetistirmeB3P1Liste[i].tcKimlikNo == lstOrmanYetistirmeB2P1Liste[j].tcKimlikNo)
                        {
                            b1p1deVar = true;
                            b1p1Index = j;
                            break;
                        }
                    }

                    //Eğer hem kesmede hem teorikte var ise 
                    if (teorikteVar && b1p1deVar && b2p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);

                        //Her iki sınavın arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }

                    else if (teorikteVar && b1p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);

                        //Her iki sınavın arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while (((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString())) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        geciciKopya.Clear();
                    }


                    else if (teorikteVar && b2p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);

                        //Her iki sınavın arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while (((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString())) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        geciciKopya.Clear();
                    }

                    else if (b1p1deVar && b2p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);

                        //Her iki sınavın arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }

                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Her iki sınavın arasında 2 saat ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while (((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || (tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString())) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        geciciKopya.Clear();
                    }

                    else if (teorikteVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);
                        //Sınavların arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }
                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeA1T1Liste[teorikIndex].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }

                    else if (b1p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);
                        //Sınavların arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }
                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB1P1Liste[b1p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }

                    else if (b2p1deVar)
                    {
                        geciciKopya = KopyasiniAl(lstOrmanYetistirmeB3P1Liste);
                        //Sınavların arasında 2 saat oluncaya kadar dön
                        while ((geciciKopya[sayac].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()))
                        {
                            Swap(geciciKopya, sayac, sayac + 1);
                            sayac = (sayac + 1) % geciciKopya.Count;
                        }
                        //Eğer girmeye çalıştığımız index boş ise:
                        if (tempB3P1[sayac].tcKimlikNo == "")
                        {
                            tempB3P1[sayac].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }
                        //Değil ise
                        else
                        {
                            int sayacSayaci = sayac;
                            //Saatler uygun ve girmeye çalıştığımız index boş oluncaya kadar dön
                            while ((tempB3P1[sayacSayaci].girisSaati - lstOrmanYetistirmeB2P1Liste[b2p1Index].girisSaati).Duration() <= TimeSpan.Parse(sinavAralik.ToShortTimeString()) || tempB3P1[sayacSayaci].tcKimlikNo != "")
                            {
                                sayacSayaci = (sayacSayaci + 1) % tempB3P1.Count;
                            }

                            tempB3P1[sayacSayaci].SaatsizDegerleriAl(geciciKopya[sayac]);
                        }

                        geciciKopya.Clear();
                    }

                    //Eğer diğer listelerde adı yok ise eşleşemeyenler listesine ekleniyor.
                    else
                    {
                        eslesemeyenler.Add(lstOrmanYetistirmeB3P1Liste[i]);
                    }

                }

                int degisken = 0;
                while (eslesemeyenler.Any())
                {
                    //Indexteki tc kimlik numarası eğer boş ise, bu kişileri bu yerlere ekliyoruz
                    if (tempB3P1[degisken].tcKimlikNo == "")
                    {
                        tempB3P1[degisken].SaatsizDegerleriAl(eslesemeyenler[0]);
                        eslesemeyenler.RemoveAt(0);
                    }
                    degisken = (degisken + 1) % tempB3P1.Count;
                }

                lstOrmanYetistirmeB3P1Liste = tempB3P1;
            }


        }


        /// <summary>
        /// Bu fonksiyon listenin seçilen 2 değerini, saatler hariç, değiş-tokuş etmemize yarar.
        /// </summary>
        void Swap(List<Aday> list, int indexA, int indexB)
        {
            indexA %= list.Count;
            indexB %= list.Count;
            Aday tmp = list[indexA];
            DateTime time;
            time = tmp.girisSaati;
            list[indexA].girisSaati = list[indexB].girisSaati;
            list[indexB].girisSaati = time;
            list[indexA] = list[indexB];
            list[indexB] = tmp;
        }

        /// <summary>
        /// Bu fonksiyon seçilen listenin bir kopyasını, yeni nesneler üreterek oluşturur.
        /// </summary>
        List<Aday> KopyasiniAl(List<Aday> liste)
        {
            List<Aday> temp = new List<Aday>();
            for(int i = 0; i < liste.Count; i++)
            {
                Aday aday = new Aday(sinavSuresi);
                aday.DegerleriAl(liste[i]);
                temp.Add(aday);
            }
            return temp;
        }


        /// <summary>
        /// Bu fonksiyon girilen sayılar arasında en uzun sayıyı döndürmeyi sağlar
        /// </summary>
        public int EnUzunListeUzunlugu(params int[] ints)
        {
            int largest = 0;
            foreach(var sayi in ints)
            {
                if(sayi > largest)
                {
                    largest = sayi;
                }
            }
            return largest;
        }


        XLWorkbook workbook;
        /// <summary>
        /// Bu fonksiyon çıktı olarak istenen excel dosyasının hazırlanmasını sağlar.
        /// </summary>
        void OrmanUretimCetvelHazirla(string tarih)
        {

            sinavYeri = StringFromDatabase("SELECT sinavAdi FROM[One$] WHERE sinavTarihi='" + tarih + "'");
            string sinavId = StringFromDatabase("SELECT myksinavid FROM [Three$]");
            //Listelerin atamaları yapılıyor
            lstOrmanUretimA1B1B2T1Liste = DeleteDuplicates("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'T1' AND (birimkodu LIKE '%A1' OR birimkodu LIKE '%B1' OR birimkodu LIKE '%B2') AND sinavtarihi = '" + tarih + "'");
            lstOrmanUretimB1P1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B1' AND sinavtarihi = '" + tarih + "'");
            lstOrmanUretimB2P1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B2' AND sinavtarihi = '" + tarih + "'");

            if (lstOrmanUretimA1B1B2T1Liste.Count < 16 && lstOrmanUretimA1B1B2T1Liste.Any())
            {
                int count = 16 - lstOrmanUretimA1B1B2T1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanUretimA1B1B2T1Liste.Add(temp);
                }
            }

            if (lstOrmanUretimB1P1Liste.Count < 16 && lstOrmanUretimB1P1Liste.Any())
            {
                int count = 16 - lstOrmanUretimB1P1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanUretimB1P1Liste.Add(temp);
                }
            }
            if (lstOrmanUretimB2P1Liste.Count < 16 && lstOrmanUretimB2P1Liste.Any())
            {
                int count = 16 - lstOrmanUretimB2P1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanUretimB2P1Liste.Add(temp);
                }
            }

            foreach(var item in lstOrmanUretimA1B1B2T1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "A1B1B2T1";
            }
            foreach (var item in lstOrmanUretimB1P1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "B1P1";
            }
            foreach (var item in lstOrmanUretimB2P1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "B2P1";
            }

            //Teorik sınavlarının her biri kendine has bir süreye sahip, bunun ayarlaması bu fonksiyon yardımıyla yapılıyor.
            TeorikSaatveIsimAyarlama(tarih);
            //Listelerin saatleri ayarlanmadan önce, listeler kaydırılıyor (İlk adaylar için kolaylık olsun diye)
            lstOrmanUretimB1P1Liste = Rotate(lstOrmanUretimB1P1Liste, 4);
            lstOrmanUretimB2P1Liste = Rotate(lstOrmanUretimB2P1Liste, 8);

            

            //Sınav giriş süreleri ayarlanıyor.
            GirisSaatiAyarlama(lstOrmanUretimA1B1B2T1Liste);
            GirisSaatiAyarlama(lstOrmanUretimB1P1Liste);
            GirisSaatiAyarlama(lstOrmanUretimB2P1Liste);
            
            //Kişilerin sınava giriş saatleri, kişiler çakışmayacak ve aralarında en az kullanıcının seçtiği saat kadar aralık olacak şekilde ayarlanıyor.
            OrmanUretimGruplayici();

            workbook.AddWorksheet(tarih);
            var ws = workbook.Worksheet(tarih);
            //Sınav yeri, günü ve sınav id'sinin gösterildiği alan
            int row = 1;
            ws.Cell("A" + row.ToString()).Value = sinavYeri + " " + DateTime.Parse(tarih).ToString("dddd").ToUpper() + " SINAV ID: " + sinavId;
            var range = ws.Range("A1:L1");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(0, 176, 240);

            //Başlıkların ayarlanması
            row = 2;
            ws.Cell("A" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            ws.Cell("B" + row.ToString()).Value = "A1T1, B1T1 VE B2T1 - TEORİK";
            range = ws.Range("B2:C2");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("D" + row.ToString()).Value = "DEĞERLENDİRİCİ";
            ws.Cell("D" + row.ToString()).Style.Font.SetBold();
            ws.Cell("D" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            ws.Cell("D" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("E" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            ws.Cell("F" + row.ToString()).Value = "B1P1 - KESİM";
            range = ws.Range("F2:G2");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("H" + row.ToString()).Value = "DEĞERLENDİRİCİ";
            ws.Cell("H" + row.ToString()).Style.Font.SetBold();
            ws.Cell("H" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            ws.Cell("H" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("I" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            ws.Cell("J" + row.ToString()).Value = "B2P1 - SÜRÜTME";
            range = ws.Range("J2:K2");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("L" + row.ToString()).Value = "DEĞERLENDİRİCİ";
            ws.Cell("L" + row.ToString()).Style.Font.SetBold();
            ws.Cell("L" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            ws.Cell("L" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);


            //Teorik Sınavların Planlama Cetveline Eklenmesi
            row = 3;

            foreach (var item in lstOrmanUretimA1B1B2T1Liste)
            {
                ws.Cell("A" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("A" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("B" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("B" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("C" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("C" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("D" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("D" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            //Kesme Performans Sınavının Planlama Cetveline Eklenmesi

            row = 3;
            foreach (var item in lstOrmanUretimB1P1Liste)
            {
                ws.Cell("E" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("E" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("F" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("F" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("G" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("G" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("H" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("H" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            //Sürütme Performans Sınavının Planlama Cetveline Eklenmesi
            row = 3;
            foreach (var item in lstOrmanUretimB2P1Liste)
            {
                ws.Cell("I" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("I" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("J" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("J" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("K" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("K" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("L" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("L" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            var satir = ws.Rows();
            satir.Height = 35;
            satir.Style.Font.SetFontSize(12);
            satir.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            
            //Excel dosyasının alt kısmı
            row = EnUzunListeUzunlugu(lstOrmanUretimA1B1B2T1Liste.Count,lstOrmanUretimB1P1Liste.Count,lstOrmanUretimB2P1Liste.Count)+4;
            if(planlamaCetveliSatirSayisi < row - 4)
                planlamaCetveliSatirSayisi = row - 4;

            ws.Cell("A" + row.ToString()).Value = "SABAH SAAT TAM 08:00'DE SINAVA BAŞLAYACAK ŞEKİLDE SINAV ALANINI KURUNUZ.";
            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row++;
            ws.Cell("A" + row.ToString()).Value = "Kameramanlar: ";
            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row++;
            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row = row + 3;
            ws.Cell("A" + row.ToString()).Value = "BU SINAVDA YOKLAMA LİSTESİ UYGULAMASI BULUNMAMAKTADIR. ";

            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 25;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            row++;
            ws.Cell("A" + row.ToString()).Value = "ADAYLAR İMZALARINI ADAY BELGELERİNİN ARASINDA BULUNAN TAAHHÜTNAMEYE ATACAKLARDIR.";
            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 18;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            row++;
            ws.Cell("A" + row.ToString()).Value = " LÜTFEN UNUTMAYINIZ.";
            range = ws.Range("A" + row.ToString() + ":L" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 20;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            //Excel dosyamızın hücrelerinin yazılarını ortalıyoruz
            satir.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //Hücrelerin boyutunu içeriğe göre ayarlıyoruz
            ws.Columns("A","L").AdjustToContents();


            foreach (var item in lstOrmanUretimA1B1B2T1Liste)
            {
                foreach(var temp in tmpOrmanUretimA1T1full)
                {
                    if(item.tcKimlikNo == temp.tcKimlikNo)
                    {
                        temp.girisSaati = item.girisSaati;
                        temp.girecegiOturum = "A1B1B2T1";
                        temp.column = 'B';
                        temp.girisTarihi = item.girisTarihi;
                        break;
                    }
                }

                foreach (var temp in tmpOrmanUretimB1T1full)
                {
                    if (item.tcKimlikNo == temp.tcKimlikNo)
                    {
                        temp.girisSaati = item.girisSaati;
                        temp.girecegiOturum = "A1B1B2T1";
                        temp.column = 'B';
                        temp.girisTarihi = item.girisTarihi;
                        break;
                    }
                }

                foreach (var temp in tmpOrmanUretimB2T1full)
                {
                    if (item.tcKimlikNo == temp.tcKimlikNo)
                    {
                        temp.girisSaati = item.girisSaati;
                        temp.girecegiOturum = "A1B1B2T1";
                        temp.column = 'B';
                        temp.girisTarihi = item.girisTarihi;
                        break;
                    }
                }
            }

            foreach (var item in lstOrmanUretimB1P1Liste)
            {
                item.column = 'F';
                tmpOrmanUretimB1P1full.Add(item);
            }
            foreach (var item in lstOrmanUretimB2P1Liste)
            {
                item.column = 'J';
                tmpOrmanUretimB2P1full.Add(item);
            }

        }

        void OrmanYetistirmeCetvelHazirla(string tarih)
        {

            sinavYeri = StringFromDatabase("SELECT sinavAdi FROM[One$] WHERE sinavTarihi='" + tarih + "'");
            string sinavId = StringFromDatabase("SELECT myksinavid FROM [Three$]");
            //Listelerin atamaları yapılıyor
            lstOrmanYetistirmeA1T1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%A1' AND sinavtarihi = '" + tarih + "'");
            lstOrmanYetistirmeB1P1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B1' AND sinavtarihi = '" + tarih + "'");
            lstOrmanYetistirmeB2P1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B2' AND sinavtarihi = '" + tarih + "'");
            lstOrmanYetistirmeB3P1Liste = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B3' AND sinavtarihi = '" + tarih + "'");

            //Listelerin saatleri ayarlanmadan önce, listeler kaydırılıyor (İlk adaylar için kolaylık olsun diye)
            lstOrmanYetistirmeB1P1Liste = Rotate(lstOrmanYetistirmeB1P1Liste, 4);
            lstOrmanYetistirmeB2P1Liste = Rotate(lstOrmanYetistirmeB2P1Liste, 8);
            lstOrmanYetistirmeB3P1Liste = Rotate(lstOrmanYetistirmeB3P1Liste, 12);


            if (lstOrmanYetistirmeA1T1Liste.Count < 16 && lstOrmanYetistirmeA1T1Liste.Any())
            {
                int count = 16 - lstOrmanYetistirmeA1T1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanYetistirmeA1T1Liste.Add(temp);
                }
            }

            if (lstOrmanYetistirmeB1P1Liste.Count < 16 && lstOrmanYetistirmeB1P1Liste.Any())
            {
                int count = 16 - lstOrmanYetistirmeB1P1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanYetistirmeB1P1Liste.Add(temp);
                }
            }
            if (lstOrmanYetistirmeB2P1Liste.Count < 16 && lstOrmanYetistirmeB2P1Liste.Any())
            {
                int count = 16 - lstOrmanYetistirmeB2P1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanYetistirmeB2P1Liste.Add(temp);
                }
            }
            if (lstOrmanYetistirmeB3P1Liste.Count < 16 && lstOrmanYetistirmeB3P1Liste.Any())
            {
                int count = 16 - lstOrmanYetistirmeB3P1Liste.Count;
                for (int i = 0; i < count; i++)
                {
                    Aday temp = new Aday(sinavSuresi);
                    lstOrmanYetistirmeB3P1Liste.Add(temp);
                }
            }


            foreach (var item in lstOrmanYetistirmeA1T1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "A1T1";
            }
            foreach (var item in lstOrmanYetistirmeB1P1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "B1P1";
            }
            foreach (var item in lstOrmanYetistirmeB2P1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "B2P1";
            }
            foreach (var item in lstOrmanYetistirmeB3P1Liste)
            {
                item.girisTarihi = tarih;
                item.girecegiOturum = "B3P1";
            }

            //Sınav giriş süreleri ayarlanıyor.
            GirisSaatiAyarlama(lstOrmanYetistirmeA1T1Liste);
            GirisSaatiAyarlama(lstOrmanYetistirmeB1P1Liste);
            GirisSaatiAyarlama(lstOrmanYetistirmeB2P1Liste);
            GirisSaatiAyarlama(lstOrmanYetistirmeB3P1Liste);

            //Kişilerin sınava giriş saatleri, kişiler çakışmayacak ve aralarında en az kullanıcının seçtiği saat kadar aralık olacak şekilde ayarlanıyor.
            OrmanYetistirmeGruplayici();

            workbook.AddWorksheet(tarih);
            var ws = workbook.Worksheet(tarih);
            //Sınav yeri, günü ve sınav id'sinin gösterildiği alan
            int row = 1;
            ws.Cell("A" + row.ToString()).Value = sinavYeri + " " + DateTime.Parse(tarih).ToString("dddd").ToUpper() + " SINAV ID: " + sinavId;
            var range = ws.Range("A1:P1");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(0, 176, 240);

            //Başlıkların ayarlanması
            row = 2;
            ws.Cell("A" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            ws.Cell("B" + row.ToString()).Value = "A1T1 - TEORİK";
            range = ws.Range("B2:C2");
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("D" + row.ToString()).Value = "DEĞERLENDİRİCİ";
            ws.Cell("D" + row.ToString()).Style.Font.SetBold();
            ws.Cell("D" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            ws.Cell("D" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            ws.Cell("E" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);

            if (lstOrmanYetistirmeB1P1Liste.Any())
            {
                ws.Cell("F" + row.ToString()).Value = "B1P1 - AĞAÇLANDIRMA";
                range = ws.Range("F2:G2");
                range.Merge().Style.Font.SetBold();
                range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

                ws.Cell("H" + row.ToString()).Value = "DEĞERLENDİRİCİ";
                ws.Cell("H" + row.ToString()).Style.Font.SetBold();
                ws.Cell("H" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("H" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
            }

            if (lstOrmanYetistirmeB2P1Liste.Any())
            {
                ws.Cell("I" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("J" + row.ToString()).Value = "B2P1 - FİDAN YETİŞTİRME";
                range = ws.Range("J2:K2");
                range.Merge().Style.Font.SetBold();
                range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

                ws.Cell("L" + row.ToString()).Value = "DEĞERLENDİRİCİ";
                ws.Cell("L" + row.ToString()).Style.Font.SetBold();
                ws.Cell("L" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("L" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
            }

            

            if (lstOrmanYetistirmeB3P1Liste.Any())
            {
                ws.Cell("M" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("N" + row.ToString()).Value = "B3P1 - AŞI VE ÇELİKLE ÜRETİM";
                range = ws.Range("N2:O2");
                range.Merge().Style.Font.SetBold();
                range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

                ws.Cell("P" + row.ToString()).Value = "DEĞERLENDİRİCİ";
                ws.Cell("P" + row.ToString()).Style.Font.SetBold();
                ws.Cell("P" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("P" + row.ToString()).Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);
            }


            //Teorik Sınavların Planlama Cetveline Eklenmesi
            row = 3;

            foreach (var item in lstOrmanYetistirmeA1T1Liste)
            {
                ws.Cell("A" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("A" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("B" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("B" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("C" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("C" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("D" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("D" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            //Kesme Performans Sınavının Planlama Cetveline Eklenmesi

            row = 3;
            foreach (var item in lstOrmanYetistirmeB1P1Liste)
            {
                ws.Cell("E" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("E" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("F" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("F" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("G" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("G" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("H" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("H" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            //Sürütme Performans Sınavının Planlama Cetveline Eklenmesi
            row = 3;
            foreach (var item in lstOrmanYetistirmeB2P1Liste)
            {
                ws.Cell("I" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("I" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("J" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("J" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("K" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("K" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("L" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("L" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            row = 3;
            foreach (var item in lstOrmanYetistirmeB3P1Liste)
            {
                ws.Cell("M" + row.ToString()).SetValue(item.girisSaati.ToShortTimeString());
                ws.Cell("M" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("N" + row.ToString()).Value = item.tcKimlikNo;
                ws.Cell("N" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("O" + row.ToString()).Value = item.adiSoyadi;
                ws.Cell("O" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                ws.Cell("P" + row.ToString()).Value = item.degerlendirici;
                ws.Cell("P" + row.ToString()).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                row++;
            }

            var satir = ws.Rows();
            satir.Height = 35;
            satir.Style.Font.SetFontSize(12);
            satir.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

            //Excel dosyasının alt kısmı
            row = EnUzunListeUzunlugu(lstOrmanYetistirmeA1T1Liste.Count,lstOrmanYetistirmeB1P1Liste.Count,lstOrmanYetistirmeB2P1Liste.Count,lstOrmanYetistirmeB3P1Liste.Count) + 4;
            planlamaCetveliSatirSayisi = row-4;

            ws.Cell("A" + row.ToString()).Value = "SABAH SAAT TAM 08:00'DE SINAVA BAŞLAYACAK ŞEKİLDE SINAV ALANINI KURUNUZ.";
            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.SetBold().Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row++;
            ws.Cell("A" + row.ToString()).Value = "Kameramanlar: ";
            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row++;
            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 14;
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(191, 191, 191);

            row = row + 3;
            ws.Cell("A" + row.ToString()).Value = "BU SINAVDA YOKLAMA LİSTESİ UYGULAMASI BULUNMAMAKTADIR. ";

            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 25;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            row++;
            ws.Cell("A" + row.ToString()).Value = "ADAYLAR İMZALARINI ADAY BELGELERİNİN ARASINDA BULUNAN TAAHHÜTNAMEYE ATACAKLARDIR.";
            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 18;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            row++;
            ws.Cell("A" + row.ToString()).Value = " LÜTFEN UNUTMAYINIZ.";
            range = ws.Range("A" + row.ToString() + ":P" + row.ToString());
            range.Merge().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Merge().Style.Font.FontSize = 20;
            range.Merge().Style.Font.SetBold();
            range.Merge().Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            range.Merge().Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 0);

            //Excel dosyamızın hücrelerinin yazılarını ortalıyoruz
            satir.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            //Hücrelerin boyutunu içeriğe göre ayarlıyoruz
            ws.Columns("A", "P").AdjustToContents();

            foreach (var item in lstOrmanYetistirmeA1T1Liste)
            {
                item.column = 'B';
                item.girisTarihi = tarih;
                item.girecegiOturum = "A1T1";
                tmpOrmanYetistirmeA1T1full.Add(item);
            }
            foreach (var item in lstOrmanYetistirmeB1P1Liste)
            {

                item.column = 'F';
                item.girisTarihi = tarih;
                item.girecegiOturum = "B1P1";
                tmpOrmanYetistirmeB1P1full.Add(item);
            }
            foreach (var item in lstOrmanYetistirmeB2P1Liste)
            {

                item.column = 'J';
                item.girisTarihi = tarih;
                item.girecegiOturum = "B2P1";
                tmpOrmanYetistirmeB2P1full.Add(item);
            }
            foreach (var item in lstOrmanYetistirmeB3P1Liste)
            {

                item.column = 'N';
                item.girisTarihi = tarih;
                item.girecegiOturum = "B3P1";
                tmpOrmanYetistirmeB3P1full.Add(item);
            }
        }

        /// <summary>
        /// Bu fonksiyon oluşturduğumuz çizelgenin bir json dosyası halinde döndürülmesini sağlar
        /// </summary>
        void OrmanUretimJsonlastirma()
        {
            List<Liste> liste = new List<Liste>();
            foreach (var item in tmpOrmanUretimA1T1full)
            {
                if(item.tcKimlikNo != "")
                {
                    item.satir = OrmanUretimSatirSayi(item.tcKimlikNo, "teorikA1");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()
                    };
                    liste.Add(eleman);
                }
            }
            foreach (var item in tmpOrmanUretimB1T1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanUretimSatirSayi(item.tcKimlikNo, "teorikB1");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()
                    };
                    liste.Add(eleman);
                }
            }
            foreach (var item in tmpOrmanUretimB2T1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanUretimSatirSayi(item.tcKimlikNo, "teorikB2");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()
                    };
                    liste.Add(eleman);
                }
            }
            foreach (var item in tmpOrmanUretimB1P1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanUretimSatirSayi(item.tcKimlikNo, "kesme");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()

                    };
                    liste.Add(eleman);
                }
            }
            foreach (var item in tmpOrmanUretimB2P1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanUretimSatirSayi(item.tcKimlikNo, "surtme");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()

                    };
                    liste.Add(eleman);
                }
            }

            string json = JsonConvert.SerializeObject(liste);

        }

        /// <summary>
        /// Bu fonksiyon oluşturduğumuz çizelgenin bir json dosyası halinde döndürülmesini sağlar
        /// </summary>
        void OrmanYetistirmeJsonlastirma()
        {
            List<Liste> liste = new List<Liste>();
            foreach (var item in tmpOrmanYetistirmeA1T1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanYetistirmeSatirSayi(item.tcKimlikNo, "teorikA1");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()
                    };
                    liste.Add(eleman);
                }
            }

            foreach (var item in tmpOrmanYetistirmeB1P1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanYetistirmeSatirSayi(item.tcKimlikNo, "performansB1");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()

                    };
                    liste.Add(eleman);
                }
            }

            foreach (var item in tmpOrmanYetistirmeB2P1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanYetistirmeSatirSayi(item.tcKimlikNo, "performansB2");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()

                    };
                    liste.Add(eleman);
                }
            }

            foreach (var item in tmpOrmanYetistirmeB3P1full)
            {
                if (item.tcKimlikNo != "")
                {
                    item.satir = OrmanYetistirmeSatirSayi(item.tcKimlikNo, "performansB3");
                    Liste eleman = new Liste()
                    {
                        to_basvuru_sinavQid = IdDondur(item.satir),
                        sinavsaati = item.girisSaati.ToShortTimeString()

                    };
                    liste.Add(eleman);
                }
            }

            string json = JsonConvert.SerializeObject(liste);

        }
        
        /// <summary>
        /// Bu fonksiyon satır sayısını aldığımız değerin olduğu kısımdaki id değerini almamızı sağlar
        /// </summary>
        /// <param name="row">Satır</param>
        /// <returns></returns>
        public string IdDondur(int row)
        {
            XLWorkbook wb = new XLWorkbook(dosyaYolu);
            var ws = wb.Worksheet("Four");
            string id = ws.Cell("A" + row.ToString()).Value.ToString();
            return id;
        }

        /// <summary>
        /// Bu fonksiyon verdiğimiz tc kimlik numaralı kişinin sınavına göre aday bildirim dosyasındaki satır sayısını almamızı sağlar.
        /// </summary>
        /// <param name="tckn">TC Kimlik Numarası</param>
        /// <param name="sinavTuru">Sınav Türü</param>
        /// <returns></returns>
        /// 
        public int OrmanUretimSatirSayi(string tckn, string sinavTuru)
        {
            int rowIndex = 0;
            switch (sinavTuru)
            {
                case "teorikA1":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("A1") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("T1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;


                case "teorikB1":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B1") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("T1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;


                case "teorikB2":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B2") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("T1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;

                case "kesme":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B1") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("P1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;

                case "surtme":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B2") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("P1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;
            }
            return rowIndex + 2;
        }

        /// <summary>
        /// Bu fonksiyon verdiğimiz tc kimlik numaralı kişinin sınavına göre aday bildirim dosyasındaki satır sayısını almamızı sağlar.
        /// </summary>
        /// <param name="tckn">TC Kimlik Numarası</param>
        /// <param name="sinavTuru">Sınav Türü</param>
        /// <returns></returns>
        /// 
        public int OrmanYetistirmeSatirSayi(string tckn, string sinavTuru)
        {
            int rowIndex = 0;
            switch (sinavTuru)
            {
                case "teorikA1":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("A1") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("T1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;


                case "performansB1":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B1") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("P1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;

                case "performansB2":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B2") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("P1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;

                case "performansB3":
                    foreach (DataGridViewRow row in ormanUretimData.Rows)
                    {
                        if (row.Cells["tckn"].Value.ToString() == tckn && row.Cells["birimkodu"].Value.ToString().EndsWith("B3") && row.Cells["sinavturukodu"].Value.ToString().EndsWith("P1"))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }
                    break;
            }
            return rowIndex + 2;
        }

        private void btnDosyaSec_Click(object sender, EventArgs e)
        {
            //Kullanıcı bu butona tıkladıktan sonra kendisine bir dosya seçme alanı gösteriliyor.
            //Kullanıcı dosyayı seçtiğinde ise path alınıyor.
            dosyaYolu = DosyaAdiGetir();
            //Connection String ayarlanıyor bu sayede bunu her yerde kullanabiliriz.
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'";
            //Bu path ile beraber dosyaya gidilip dataGridView1'e excel tablosunun içindeki veriler yazılıyor.
            ExcelSorgu("SELECT * FROM [One$]");
            //Tabloları oluştur butonu görünür yapılıyor
            webBrowser1.Visible = false;
           

        }

        private void btnOrmanUretimTabloOlustur_Click(object sender, EventArgs e)
        {

            sinavSuresi = int.Parse(txtSinavSuresi.Text);
            //Adayların 2 sınavı arasındaki saat farkını ayarlamak için kullanıcıdan girdi alıyoruz
            sinavAralik = DateTime.Parse(txtOrmanUretimSinavAralik.Text).AddMinutes(-1);
            //Eğer sinavAralik girdisi 2 saatten fazla veya yarım saatten az ise
            if(sinavAralik > DateTime.Parse("01:59") || sinavAralik < DateTime.Parse("00:29"))
            {
                MessageBox.Show("Lütfen aralığı 2 saatten fazla veya yarım saatten az yapmayın");
                return;
            }
            //Progress bar'ın maksimum ve minimum noktaları eklendi. Progress bar'ın değeri bu fonksiyon her çağrıldığında sıfırlanacak şekilde ayarlandı.
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            //Bu fonksiyon bizim yeni bir excel dosyası oluşturmamızı sağlıyor.
            workbook = new XLWorkbook();
            //Sınav başlangıç saatini kullanıcıdan alıyoruz.
            sinavBaslangic = DateTime.Parse(txtOrmanUretimBaslangicSaati.Text);
            //Excel tablosunda bulunan tarihleri bir sorgu ile liste şeklinde alıyoruz
            var tarihler = ListFromDatabase("SELECT sinavTarihi FROM [One$] GROUP BY sinavtarihi");
            //Eğer tarih formatına uymayan bir veri eklendiyse, listeden atıyoruz.
            tarihler = TarihFormat(tarihler);
            //Her bir tarih için bu döngüyü döndürüyoruz.
            foreach(var tarih in tarihler)
            {
                //Progress barı biten her işlem için yüzdelik olarak arttırıyoruz
                progressBar1.Value += 100 / tarihler.Count;
                //Belirlenen tarihin cetveli ayarlanıyor.
                OrmanUretimCetvelHazirla(tarih);
            }

            //Kaydedilecek dosyanın adını düzenliyoruz
            //Aralık değişkeni sınavın yapıldığı ilk gün ve son günü tutan bir değişken.
            string aralik = DateTime.Parse(tarihler.First()).ToString("dd/MMMM");
            aralik += "-" + DateTime.Parse(tarihler.Last()).ToString("dd/MMMM");
            //Kayıt değişkeni dosyanın kaydedileceği path'i tutan değişken
            string kayit = Path.GetDirectoryName(dosyaYolu);
            //Sınav yeri ve gününü tutan sinavYeri değişkeninin ilk kelimesini alıyoruz. Böylelikle sınav lokasyonunu elde ediyoruz
            var gecici = Regex.Match(sinavYeri, @"^([\w\-]+)");
            //Dosyamıza ismini veriyoruz.
            kayit = kayit + "\\Planlama Cetveli " + aralik + " " + gecici.Value + " .xlsx";
            //Oluşturulan excel dosyasını kaydediyoruz.
            workbook.SaveAs(kayit);
            //İşlemin bittiğini göstermek için progress bar'ı dolduruyoruz.
            progressBar1.Value = 100;
            //İşlemin bittiğini göstermek için kullanıcıya mesaj gönderiyoruz.
            MessageBox.Show("İşleminiz Başarı ile Tamamlandı.");

            DialogResult dialogResult = MessageBox.Show("Saatler sisteme girilsin mi?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                dialogResult = MessageBox.Show("Oluşturulan planlama cetvelinde saatleri güncellemek ister misiniz?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    PlanlamaCetveliForm p = new PlanlamaCetveliForm();
                    p.Olusturucu(kayit, tarihler,planlamaCetveliSatirSayisi, sinavSuresi);
                    p.Listeler(tmpOrmanUretimA1T1full, tmpOrmanUretimB1P1full, tmpOrmanUretimB1T1full, tmpOrmanUretimB2P1full, tmpOrmanUretimB2T1full);
                    p.ShowDialog();
                    p.fullAdayListesi[0] = tmpOrmanUretimA1T1full; 
                    p.fullAdayListesi[1] = tmpOrmanUretimB1P1full; 
                    p.fullAdayListesi[2] = tmpOrmanUretimB1T1full;
                    p.fullAdayListesi[3] = tmpOrmanUretimB2P1full;
                    p.fullAdayListesi[4] = tmpOrmanUretimB2T1full;

                }
                OrmanUretimJsonlastirma();
            }
            else
            {
                //Kaydedilen dosyayı, dosya gezgininde gösteriyoruz.
                Process.Start("explorer.exe", string.Format("/select,\"{0}\"", kayit));
            }

        }



        private void btnOrmanYetistirmeVeBakimTabloOlustur_Click(object sender, EventArgs e)
        {
            sinavSuresi = int.Parse(txtOrmanYetistirmeSinavSuresi.Text);
            sinavAralik = DateTime.Parse(txtOrmanYetistirmeVeBakimSinavAralik.Text).AddMinutes(-1);
            //Eğer sinavAralik girdisi 2 saatten fazla veya yarım saatten az ise
            if (sinavAralik > DateTime.Parse("01:59") || sinavAralik < DateTime.Parse("00:29"))
            {
                MessageBox.Show("Lütfen aralığı 2 saatten fazla veya yarım saatten az yapmayın");
                return;
            }
            //Progress bar'ın maksimum ve minimum noktaları eklendi. Progress bar'ın değeri bu fonksiyon her çağrıldığında sıfırlanacak şekilde ayarlandı.
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            //Bu fonksiyon bizim yeni bir excel dosyası oluşturmamızı sağlıyor.
            workbook = new XLWorkbook();
            //Sınav başlangıç saatini kullanıcıdan alıyoruz.
            sinavBaslangic = DateTime.Parse(txtOrmanYetistirmeVeBakimBaslangicSaati.Text);
            //Excel tablosunda bulunan tarihleri bir sorgu ile liste şeklinde alıyoruz
            var tarihler = ListFromDatabase("SELECT sinavTarihi FROM [One$] GROUP BY sinavtarihi");
            //Eğer tarih formatına uymayan bir veri eklendiyse, listeden atıyoruz.
            tarihler = TarihFormat(tarihler);
            //Her bir tarih için bu döngüyü döndürüyoruz.
            foreach (var tarih in tarihler)
            {
                //Progress barı biten her işlem için yüzdelik olarak arttırıyoruz
                progressBar1.Value += 100 / tarihler.Count;
                //Belirlenen tarihin cetveli ayarlanıyor.
                OrmanYetistirmeCetvelHazirla(tarih);
            }

            //Kaydedilecek dosyanın adını düzenliyoruz
            //Aralık değişkeni sınavın yapıldığı ilk gün ve son günü tutan bir değişken.
            string aralik = DateTime.Parse(tarihler.First()).ToString("dd/MMMM");
            aralik += "-" + DateTime.Parse(tarihler.Last()).ToString("dd/MMMM");
            //Kayıt değişkeni dosyanın kaydedileceği path'i tutan değişken
            string kayit = Path.GetDirectoryName(dosyaYolu);
            //Sınav yeri ve gününü tutan sinavYeri değişkeninin ilk kelimesini alıyoruz. Böylelikle sınav lokasyonunu elde ediyoruz
            var gecici = Regex.Match(sinavYeri, @"^([\w\-]+)");
            //Dosyamıza ismini veriyoruz.
            kayit = kayit + "\\Planlama Cetveli " + aralik + " " + gecici.Value + " .xlsx";
            //Oluşturulan excel dosyasını kaydediyoruz.
            workbook.SaveAs(kayit);
            //İşlemin bittiğini göstermek için progress bar'ı dolduruyoruz.
            progressBar1.Value = 100;
            //İşlemin bittiğini göstermek için kullanıcıya mesaj gönderiyoruz.
            MessageBox.Show("İşleminiz Başarı ile Tamamlandı.");
            DialogResult dialogResult = MessageBox.Show("Saatler sisteme girilsin mi?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                dialogResult = MessageBox.Show("Oluşturulan planlama cetvelinde saatleri güncellemek ister misiniz?", "", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    PlanlamaCetveliForm p = new PlanlamaCetveliForm();
                    p.Olusturucu(kayit, tarihler, planlamaCetveliSatirSayisi, sinavSuresi);
                    p.Listeler(tmpOrmanYetistirmeA1T1full, tmpOrmanYetistirmeB1P1full, tmpOrmanYetistirmeB2P1full, tmpOrmanYetistirmeB3P1full);
                    p.ShowDialog();
                    p.fullAdayListesi[0] = tmpOrmanYetistirmeA1T1full;
                    p.fullAdayListesi[1] = tmpOrmanYetistirmeB1P1full;
                    p.fullAdayListesi[2] = tmpOrmanYetistirmeB2P1full;
                    p.fullAdayListesi[3] = tmpOrmanYetistirmeB3P1full;

                }
                OrmanYetistirmeJsonlastirma();
            }

            else
            {
                //Kaydedilen dosyayı, dosya gezgininde gösteriyoruz.
                Process.Start("explorer.exe", string.Format("/select,\"{0}\"", kayit));
            }

            //Kaydedilen dosyayı, dosya gezgininde gösteriyoruz.
            Process.Start("explorer.exe", string.Format("/select,\"{0}\"", kayit));
        }

        
        /// <summary>
        /// Bu fonksiyon datagridview'e sürükleme işlemi yapıldığında dosyanın path'ini alır ve sorgu işlemini yapar
        /// </summary>
        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                string[] deneme = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (var item in deneme)
                {
                    dosyaYolu = item;
                }

                connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'";
                ExcelSorgu("SELECT * FROM  [One$]");
                webBrowser1.Visible = false;
            }
            catch { }

        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            //Güncelleme sınıfını, çekeceğimiz json formatını class yapısına çevirmek için oluşturuyoruz
            Guncelleme gunc;
            //versiyon değişkenimiz, sürümümüzün güncel olup olmadığını kontrol edecek
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            //Sürümümüzü uygulamamızın alt tarafına yerleştiriyoruz
            lblVersion.Text = version;
            //Çektiğimiz json dosyası string formatında alınıyor
            var jsonDosyasi = new WebClient().DownloadString("http://sayemyis.subu.edu.tr/third-party-software/3/json");
            //Çektiğimiz json stringi, Güncelleme classından bir objeye çevriliyor
            gunc = JsonConvert.DeserializeObject<Guncelleme>(jsonDosyasi);
            //versiyonlarımızı kontrol etmenin çok rahat bir yolu olarak, iki versiyon değişkenimizi versiyon sınıfından bir obje olarak alıyoruz.
            var updatedVersion = new Version(gunc.version);
            var currentVersion = new Version(version);
            //Eğer güncel değilsek
            if (updatedVersion != currentVersion)
            {
                //Güncelleme işlemleri
                MessageBox.Show("Bir güncelleme var! Lütfen güncellemenin yüklenmesini bekleyin.");
                GuncellemeYap(gunc);
                Application.Exit();
            }
            else if (!gunc.status)
            {
                //Eğer uygulamamız bakımdaysa
                MessageBox.Show(gunc.statusComment);
                Application.Exit();
            }
            //webBrowser'ın güncelleme notlarını çekmesi için gittiği adres
            webBrowser1.Navigate("http://sayemyis.subu.edu.tr/third-party-software/3/updatenotes");

        }

        /// <summary>
        /// Bu fonksiyon uygulamaya güncelleme yapmamızı sağlar
        /// </summary>
        /// <param name="gunc">Guncelleme sınıfından bir obje</param>
        private void GuncellemeYap(Guncelleme gunc)
        {
            //Güncelleme dosyalarımızı yükleyeceğimiz path.
            string guncellemePath = @"C:\SinavPlanlamaCetveli\guncelleme";
            //Eğer böyle bir klasör varsa, çakışmaması adına siliyoruz.
            if (Directory.Exists(guncellemePath)) Directory.Delete(guncellemePath, true);

            using (var client = new WebClient())
            {
                if (!Directory.Exists(guncellemePath))
                {
                    //Klasörü tekrar oluşturuyoruz.
                    Directory.CreateDirectory(guncellemePath);
                }
                //Güncellememizi indiriyoruz.
                client.DownloadFile(gunc.path, guncellemePath + "\\" + gunc.version + ".zip");
            }
            //İndirdiğimiz zip dosyasını çıkartıyoruz.
            ZipFile.ExtractToDirectory(guncellemePath + "\\" + gunc.version + ".zip", guncellemePath);
            //Çıkarttığımız setup dosyasını açıyoruz ve kurulum dosyası açılıyor.
            Process.Start(guncellemePath + @"\setup.exe");
        }

        private void btnOrmanUretimPlanlamaCetvelindenDuzenle_Click(object sender, EventArgs e)
        {
            sinavSuresi = int.Parse(txtSinavSuresi.Text);
            tmpOrmanUretimA1T1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%A1'" ,true);
            tmpOrmanUretimB1T1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%B1'", true);
            tmpOrmanUretimB2T1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%B2'", true);
            tmpOrmanUretimB1P1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B1'", true);
            tmpOrmanUretimB2P1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B2'", true);
            var tarihler = ListFromDatabase("SELECT sinavTarihi FROM [One$] GROUP BY sinavtarihi");
            tarihler = TarihFormat(tarihler);

            foreach (var item in tmpOrmanUretimA1T1full)
            {
                item.girecegiOturum = "A1B1B2T1";
                item.column = 'B';
            }
            foreach (var item in tmpOrmanUretimB1T1full)
            {
                item.girecegiOturum = "A1B1B2T1";
                item.column = 'B';
            }
            foreach (var item in tmpOrmanUretimB2T1full)
            {
                item.girecegiOturum = "A1B1B2T1";
                item.column = 'B';
            }
            foreach (var item in tmpOrmanUretimB1P1full)
            {
                item.girecegiOturum = "B1P1";
                item.column = 'F';
            }
            foreach (var item in tmpOrmanUretimB2P1full)
            {
                item.girecegiOturum = "B2P1";
                item.column = 'J';
            }

            PlanlamaCetveliForm p = new PlanlamaCetveliForm();
            p.Olusturucu(null, tarihler, planlamaCetveliSatirSayisi, sinavSuresi);
            p.Listeler(tmpOrmanUretimA1T1full, tmpOrmanUretimB1P1full, tmpOrmanUretimB1T1full, tmpOrmanUretimB2P1full, tmpOrmanUretimB2T1full);
            p.ShowDialog();
            p.fullAdayListesi[0] = tmpOrmanUretimA1T1full;
            p.fullAdayListesi[1] = tmpOrmanUretimB1P1full;
            p.fullAdayListesi[2] = tmpOrmanUretimB1T1full;
            p.fullAdayListesi[3] = tmpOrmanUretimB2P1full;
            p.fullAdayListesi[4] = tmpOrmanUretimB2T1full;


            
            OrmanUretimJsonlastirma();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tmpOrmanYetistirmeA1T1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'T1' AND birimkodu LIKE '%A1'", true);
            tmpOrmanYetistirmeB1P1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B1'", true);
            tmpOrmanYetistirmeB2P1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B2'", true);
            tmpOrmanYetistirmeB3P1full = SinavFromDatabase("SELECT tckn, adi + ' ' + soyadi AS kisi, degerlendirici, sinavtarihi FROM [One$] WHERE sinavturukodu = 'P1' AND birimkodu LIKE '%B3'", true);
            var tarihler = ListFromDatabase("SELECT sinavTarihi FROM [One$] GROUP BY sinavtarihi");
            tarihler = TarihFormat(tarihler);

            foreach (var item in tmpOrmanYetistirmeA1T1full)
            {
                item.girecegiOturum = "A1T1";
                item.column = 'B';
            }
            foreach (var item in tmpOrmanYetistirmeB1P1full)
            {
                item.girecegiOturum = "B1P1";
                item.column = 'F';
            }
            foreach (var item in tmpOrmanYetistirmeB2P1full)
            {
                item.girecegiOturum = "B2P1";
                item.column = 'J';
            }
            foreach (var item in tmpOrmanYetistirmeB3P1full)
            {
                item.girecegiOturum = "B3P1";
                item.column = 'N';
            }

            PlanlamaCetveliForm p = new PlanlamaCetveliForm();
            p.Olusturucu(null, tarihler, planlamaCetveliSatirSayisi, sinavSuresi);
            p.Listeler(tmpOrmanYetistirmeA1T1full, tmpOrmanYetistirmeB1P1full, tmpOrmanYetistirmeB2P1full, tmpOrmanYetistirmeB3P1full);
            p.ShowDialog();
            p.fullAdayListesi[0] = tmpOrmanYetistirmeA1T1full;
            p.fullAdayListesi[1] = tmpOrmanYetistirmeB1P1full;
            p.fullAdayListesi[2] = tmpOrmanYetistirmeB2P1full;
            p.fullAdayListesi[3] = tmpOrmanYetistirmeB3P1full;
            OrmanYetistirmeJsonlastirma();

        }
    }
}
