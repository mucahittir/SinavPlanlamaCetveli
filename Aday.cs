using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sınav_Planlama_Cetveli
{
    public class Aday
    {
        public string tcKimlikNo;
        public string adiSoyadi;
        public string degerlendirici;
        public DateTime girisSaati;
        public int sinavSuresi;
        public int satir;
        public string girisTarihi;
        public string girecegiOturum;
        public char column;

        public Aday(int sinavSuresi)
        {
            tcKimlikNo = "";
            adiSoyadi = "";
            degerlendirici = "";
            this.sinavSuresi = sinavSuresi;
            satir = 0;
            girisTarihi = "";
            girecegiOturum = "";
            column = ' ';
            
        }

        public void DegerleriAl(Aday aday)
        {
            tcKimlikNo = aday.tcKimlikNo;
            adiSoyadi = aday.adiSoyadi;
            degerlendirici = aday.degerlendirici;
            sinavSuresi = aday.sinavSuresi;
            girisSaati = aday.girisSaati;
            satir = aday.satir;
            girisTarihi = aday.girisTarihi;
            girecegiOturum = aday.girecegiOturum;
            column = aday.column;
        }
        public void SaatsizDegerleriAl(Aday aday)
        {
            tcKimlikNo = aday.tcKimlikNo;
            adiSoyadi = aday.adiSoyadi;
            degerlendirici = aday.degerlendirici;
            sinavSuresi = aday.sinavSuresi;
            satir = aday.satir;
            girisTarihi = aday.girisTarihi;
            girecegiOturum = aday.girecegiOturum;
            column = aday.column;

        }
    }

}
