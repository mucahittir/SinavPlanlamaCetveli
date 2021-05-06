using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sinav_Planlama_Cetveli
{

    public class Rootobject
    {
        public Webhookresult webhookResult { get; set; }
        public Ki_Kullanicilarresult ki_kullanicilarResult { get; set; }
    }

    public class Webhookresult
    {
        public int id { get; set; }
        public string value { get; set; }
        public string created_at { get; set; }
        public int status { get; set; }
        public string updated_at { get; set; }
        public int available { get; set; }
    }

    public class Ki_Kullanicilarresult
    {
        public int id { get; set; }
        public string kadi { get; set; }
        public string sifre { get; set; }
        public string eposta { get; set; }
        public string telefon { get; set; }
        public object unvan { get; set; }
        public string adi { get; set; }
        public string soyadi { get; set; }
        public int delete { get; set; }
        public int roller_id { get; set; }
        public string created_at { get; set; }
        public string updated_at { get; set; }
        public string image { get; set; }
        public string tcno { get; set; }
        public string password { get; set; }
        public string email { get; set; }
        public string name { get; set; }
        public object remember_token { get; set; }
        public object onesignalId { get; set; }
        public int ntf_status { get; set; }
        public object onayCode { get; set; }
        public string last_session { get; set; }
        public int tokenQid { get; set; }
    }



}
