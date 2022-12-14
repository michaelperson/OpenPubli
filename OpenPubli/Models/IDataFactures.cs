using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace OpenPubli.Models
{
    public interface IDataFactures
    {
        [JsonPropertyOrder(1)]
        [JsonPropertyName("Description")]
        string DESCRIPTION { get; set; }
        [JsonPropertyOrder(2)]
        [JsonPropertyName("Prix Unitaire (HT)")]
        
        float PRIX_UNITAIRE_HT { get; set; }
        [JsonPropertyOrder(4)]
        [JsonPropertyName("Quantité")]
        float QUANTITE { get; set; }
        [JsonPropertyOrder(0)]

        [JsonPropertyName("Référence")]
        string REF { get; set; }
        [JsonPropertyOrder(3)]
        [JsonPropertyName("Taux Tva")]
        int TAUX_TVA { get; set; } 
    }
}