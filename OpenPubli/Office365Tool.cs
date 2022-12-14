using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using static Microsoft.Graph.Constants;

namespace OpenPubli
{
    /// <summary>
    /// Class permettant de sauvegarder et convertir des document word au travers d'office365 et de sharepoint onlines
    /// </summary>
    public  class Office365Tool
    {
        private readonly string _appId;
        private readonly string _appSecret;
        private readonly string _tenantId;
        /// <summary>
        /// Constructeur de l'outils permettant de transmettre les paramètres pour les connexion Office365
        /// </summary>
        /// <param name="appId">l'application Id disponible sous Azure active directory --> Application d'entreprise</param>
        /// <param name="appSecret">L'app secret créé dans Azure Active Directory --> Inscription d'application --> certificat & secrets</param>
        /// <param name="tenantId">L'id du client récupéré dans Azure Active directory --> vue d'ensemble</param>
        public Office365Tool(string appId, string appSecret, string tenantId)
        {
            this._appId=appId;
            this._appSecret=appSecret;
            this._tenantId = tenantId;
        }
        
        /// <summary>
        /// methode permettant de sauvegarde le document docx en docx et pdf dans le site sharepoint
        /// </summary>
        /// <param name="filePath">Chemin vers le fichier docx local</param>
        /// <param name="FinalName">Nom final dans sharepoint (sans l'extension)</param>
        public void SaveTodrive(string filePath, string FinalName)
        {
            GraphServiceClient graphClient = GetGraphServiceClientOnBehalf();

            IGraphServiceSitesCollectionPage sites = graphClient.Sites.Request().GetAsync().Result;
     
            
            try
            {
                var site = graphClient.Sites[sites?.FirstOrDefault()?.Id].Drives.Request().GetAsync().Result;
                Drive? Drivestore = site?.Where(s=>s.Name=="Factures Belgrain")?.FirstOrDefault();
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                   DriveItem doc = graphClient.Drives[Drivestore?.Id].Root
                                 .ItemWithPath($"{FinalName}.docx")
                                 .Content
                                 .Request()
                                 .PutAsync<DriveItem>(fileStream).Result;
                     SaveToPdf(doc, graphClient, Drivestore, FinalName);
                    
                        
                    
                }
               
                  


            }
            catch (Exception ex)
            {

                throw;
            }

        }

        /// <summary>
        /// Methode permettant de récupérer le docx en pdf et le sauvegarder dans le drive sharepoint
        /// </summary>
        /// <param name="di">un objet <see cref="DriveItem"/> contenant les informations du docx</param>
        /// <param name="graphClient">un objet de type<see cref="GraphServiceClient"/> généré par la méthode GetGraphServiceClientOnBehalf</param>
        /// <param name="documents">Le drive de destination de type <see cref="Drive"/></param>
        /// <param name="finalName">Nom final du document pdf (sans l'extension)</param>
        private void SaveToPdf(DriveItem di, GraphServiceClient graphClient, Drive? documents, string finalName)
        {
            if (documents == null) throw new ArgumentNullException("documents");
             
            try
            {
                List<QueryOption> queryOptions = new List<QueryOption>()
                {
                    new QueryOption("format", "pdf")
                };

                HttpRequestMessage pdfRequestMessage = graphClient.Drives[documents?.Id].Root
                  .ItemWithPath($"{finalName}.docx")
                  .Content
                  .Request(queryOptions).GetHttpRequestMessage();
                
                HttpClient cli = new HttpClient();
                cli.DefaultRequestHeaders.Authorization = GetAuthenticationHeaderValue();
                HttpResponseMessage response = cli.SendAsync(pdfRequestMessage).Result;
                using (Stream stream = response.Content.ReadAsStream())
                   {                          
                      DriveItem pdfDI = graphClient.Drives[documents?.Id].Root
                                       .ItemWithPath($"{finalName}.pdf")
                                       .Content
                                       .Request()
                                       .PutAsync<DriveItem>(stream).Result;                                       
                           
                   }
                   
               
            }
            catch (Exception ex)
            {

                throw;
            }


        }
               
        /// <summary>
        /// Permet de créer une <see cref="GraphServiceClient"/> en se basant sur les appId, secretId et tenantID
        /// </summary>
        /// <returns>une instance de <see cref="GraphServiceClient"/></returns>
        private GraphServiceClient GetGraphServiceClientOnBehalf()
        {
          
            return  new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
             
                            requestMessage.Headers.Authorization = GetAuthenticationHeaderValue();
            
                    }));

        }
        
        /// <summary>
        /// Permet de générer un Headr avec le JWt token de Ofice365
        /// </summary>
        /// <returns>Un header d'autorisation "bearer"</returns>
        private AuthenticationHeaderValue GetAuthenticationHeaderValue()
        {
           

            var clientCredential = new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential(_appId, _appSecret);
            var authority = $"https://login.microsoftonline.com/{_tenantId}";
            var authContext = new AuthenticationContext(authority);
#pragma warning disable CS0618 // Type or member is obsolete
            var token =  authContext?.AcquireTokenAsync("https://graph.microsoft.com/", clientCredential).Result;
#pragma warning restore CS0618 // Type or member is obsolete
            return new AuthenticationHeaderValue("bearer", token.AccessToken);
        }
    }

    
}
