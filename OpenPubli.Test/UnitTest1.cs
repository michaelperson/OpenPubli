using DocumentFormat.OpenXml.Drawing;
using Microsoft.IdentityModel.Abstractions;
using OpenPubli.Models;

namespace OpenPubli.Test
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void SimplePublication()
        {
            List <IDataFactures> ldatas = new List<IDataFactures>();
            ldatas.Add (new DataFactures() { DESCRIPTION="Test1", PRIX_UNITAIRE_HT=5.6F, QUANTITE=2, REF="01", TAUX_TVA=21 });
            ldatas.Add(new DataFactures() { DESCRIPTION = "Test2", PRIX_UNITAIRE_HT = 6, QUANTITE = 3, REF = "02", TAUX_TVA = 21 });
            ldatas.Add(new DataFactures() { DESCRIPTION = "Test3", PRIX_UNITAIRE_HT = 7, QUANTITE = 4, REF = "02", TAUX_TVA = 21 });

            Dictionary<string, string> fieldvalues = new Dictionary<string, string>();
            fieldvalues.Add("CompanyName", "Mike Corp");
            fieldvalues.Add("Adresse", "33 rue du colibri joyeux, 6200 LaBas");
            fieldvalues.Add("CommandNumber", "F254-22");
            fieldvalues.Add("client", "TechnoBel");
            fieldvalues.Add("CurrentDate", DateTime.Now.ToShortDateString());
            fieldvalues.Add("NomClient", "PluDeBiere Roger");
            fieldvalues.Add("AdresseClient", "All. des Artisans 19, 5590 Ciney");
            fieldvalues.Add("NoteFacture", "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque bibendum laoreet neque, eget rutrum nisi pellentesque vel. Quisque blandit ex purus. Sed vel ornare magna. Pellentesque dignissim sodales justo, nec feugiat ipsum tempor quis. Aliquam euismod aliquam est, in iaculis felis. Proin sagittis dui eget tincidunt posuere. Donec id ante risus. Nullam efficitur condimentum pretium. Vestibulum fermentum odio leo, in faucibus tellus placerat tempus. Integer consequat ullamcorper ligula sit amet varius. Ut malesuada mauris vitae nibh eleifend commodo a sed elit. Vestibulum varius mollis lobortis.");

          

            MergeTools mt = new MergeTools("Templates/BonCommande.dotx", "Generated/Final.docx", ";", fieldvalues, ldatas);


            Assert.DoesNotThrow(()=>mt.GenerateDocument());
        }

        [Test]
        public void SaveAsPdfTest()
        {
            var appId = "[Votre appId]";
            var appSecret = "[Votre appsecret]";
            var tenantId = "[Votre tenantId]";
            Random rnd = new Random();
            char[] chars = "abcdefghijklmnopqrstuvwxyz".ToCharArray();
            string randomString = "";

            for (int i = 0; i < 10; i++)
            {
                randomString += chars[rnd.Next(0, chars.Length)];
            }
             

            Office365Tool oft = new Office365Tool(appId,appSecret,tenantId);
            oft.SaveTodrive("Generated/Final.docx", $"Facture{randomString}");  
            
        }
    }
}