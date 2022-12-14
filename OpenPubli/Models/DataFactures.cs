using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenPubli.Models
{
    public class DataFactures : IDataFactures
    {
        public string REF { get; set; }
        public string DESCRIPTION { get; set; }
        public float QUANTITE { get; set; }
        public float PRIX_UNITAIRE_HT { get; set; }
        public int TAUX_TVA { get; set; }
    }
}
