using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Sınav_Planlama_Cetveli;

namespace Sinav_Planlama_Cetveli
{
    public partial class PlanlamaCetveliForm : Form
    {
        string kayit;
        int satir;
        public List<List<Aday>> fullAdayListesi;
        List<string> tarihler;
        XLWorkbook planlamaCetveli;
        int sinavSuresi;
        public PlanlamaCetveliForm()
        {
            InitializeComponent();
        }
        public void Olusturucu(string kayit, List<string> tarihler, int satir, int sinavSuresi)
        {
            this.kayit = kayit;
            this.tarihler = tarihler;
            this.satir = satir;
            this.sinavSuresi = sinavSuresi;
            fullAdayListesi = new List<List<Aday>>();

    }

        public void Listeler(params List<Aday>[] arrayOfAdaylarListesi)
        {
            foreach(var adaylarListesi in arrayOfAdaylarListesi)
            {
                fullAdayListesi.Add(adaylarListesi);
            }
        }
        public int SatirHesapla(IXLWorksheet ws)
        {
            int satir =0;
            for(int i = 3; i < ws.RowCount(); i++)
            {
                if (ws.Row(i).IsEmpty())
                {
                    satir = i-3;
                    break;
                }
            }
            return satir;
        }

        public void PlanlamaCetveliForm_Load(object sender, EventArgs e)
        {
            if(kayit != null)
            {
                Process.Start(kayit);
            }

            label1.Text = "Dosya Yolu: " + kayit;
        }
        public List<Aday> PlanlamaCetveliOku(string tarih, List<Aday> temp)
        {
            List<Aday> aday = new List<Aday>();
            if (temp.Any())
            {
                var ws = planlamaCetveli.Worksheet(tarih);
                if(satir == 0)
                {
                    satir = SatirHesapla(ws);
                }
                for (int i = 0; i < satir; i++)
                {
                    string tcKimlikNo = temp[0].column + (i + 3).ToString();
                    string girisSaati = ((char)(temp[0].column - 1)).ToString() + (i + 3).ToString();
                    string adiSoyadi = ((char)(temp[0].column + 1)).ToString() + (i + 3).ToString();
                    Aday foo;
                    try
                    {
                        foo = new Aday(sinavSuresi)
                        {

                            tcKimlikNo = ws.Cell(tcKimlikNo).Value.ToString(),
                            girisSaati = DateTime.Parse(ws.Cell(girisSaati).Value.ToString()),
                            girisTarihi = tarih,
                            adiSoyadi = ws.Cell(adiSoyadi).Value.ToString(),
                            girecegiOturum = temp[0].girecegiOturum
                        };
                        aday.Add(foo);
                    }
                    catch (FormatException)
                    {
                        break;
                    }
                }
            }
            return aday;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (var tarih in tarihler)
            {
                planlamaCetveli = new XLWorkbook(kayit);
                foreach (var temp in fullAdayListesi)
                {
                    List<Aday> foo = PlanlamaCetveliOku(tarih, temp);
                    foreach (var item in temp)
                    {
                        if (item.girisTarihi == tarih)
                        {
                            foreach (var planlamaCetvelindekiAday in foo)
                            {
                                
                                if (item.tcKimlikNo == planlamaCetvelindekiAday.tcKimlikNo && item.girecegiOturum == planlamaCetvelindekiAday.girecegiOturum && (item.tcKimlikNo != "" || planlamaCetvelindekiAday.tcKimlikNo != ""))
                                {
                                    if (item.girisSaati != planlamaCetvelindekiAday.girisSaati)
                                    {
                                        item.girisSaati = planlamaCetvelindekiAday.girisSaati;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            DialogResult dialogResult = MessageBox.Show("İşlemi onaylıyor musunuz?", "", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                Close();
            }
            else
            {

            }
        }

        private void PlanlamaCetveliForm_DragDrop(object sender, DragEventArgs e)
        {
            string[] deneme = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var item in deneme)
            {
                kayit = item;
            }
            label1.Text = "Dosya Yolu: " + kayit;
        }

        private void PlanlamaCetveliForm_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void dosyaSec_Click(object sender, EventArgs e)
        {
            kayit = DosyaAdiGetir();
        }

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
    }
}
