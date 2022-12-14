using DocumentFormat.OpenXml;
using dr =DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenPubli.Models;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text.Json.Serialization; 

namespace OpenPubli
{
    /// <summary>
    /// Mini outils de base permettant la création d'un document de facturation avec un publipostage vers word + une création de table avec les données de facturation
    /// </summary>
    public class MergeTools
    {
        private readonly string _templateFileName;
        private readonly string _targetFileName;
        private readonly string _fieldDelimiter;
        private readonly List<IDataFactures> _ldatas;
        private Dictionary<string, string> _fieldValues;
        /// <summary>
        /// Constructeur de l'outils de fusion
        /// </summary>
        /// <param name="templateFileName">Le chemin vers le fichier template dotx</param>
        /// <param name="targetFileName">Le chemin vers le fichier a générer en docx</param>
        /// <param name="fieldDelimeter">Le délimiter utilisé dans word si plusieurs signets se suivent</param>
        /// <param name="fieldValues">Un <see cref="Dictionary{TKey, TValue}"/> où la clé = le nom du champ a mergé dans word et la valeur , celle a inclure /></param>
        /// <param name="ldata">Une liste de <see cref="IDataFactures"/> contenant les informations de facturation</param>
        public MergeTools(string templateFileName, string targetFileName, string fieldDelimeter, Dictionary<string,string> fieldValues, List<IDataFactures> ldata)
        {
            this._templateFileName = templateFileName;
            this._targetFileName = targetFileName;
            this._fieldDelimiter = fieldDelimeter;
            this._fieldValues = fieldValues;
            this._ldatas = ldata;
        }

        /// <summary>
        /// Permet de générer le document de facturation
        /// </summary>
        /// <returns>True si le document a pu être généré</returns>
        public bool GenerateDocument()
        {
            try
            { 
                if (!File.Exists(_templateFileName)) { throw new Exception(message: "TemplateFileName (" + _templateFileName + ") does not exist"); }
                if (File.Exists(_targetFileName)) File.Delete(_targetFileName);
                    File.Copy(_templateFileName, _targetFileName);
                using (WordprocessingDocument docGenerated = WordprocessingDocument.Open(_targetFileName, true))
                {
                    docGenerated.ChangeDocumentType(WordprocessingDocumentType.Document);
                    foreach (FieldCode field in docGenerated.MainDocumentPart?.RootElement?.Descendants<FieldCode>())
                    {
                        var fieldNameStart = field.Text.LastIndexOf(_fieldDelimiter, StringComparison.Ordinal);
                        var fieldname = field.Text.Substring(fieldNameStart + _fieldDelimiter.Length).Trim();
                 
                        var fieldValue = GetMergeValue(FieldName: fieldname);
                        ReplaceMergeFieldWithText(field, fieldValue);
                    }
                     
                    InsertDataIntoTable(docGenerated);
                }
            
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }

        /// <summary>
        /// Permet d'insérer la table de facturation dans le document word. 
        /// L'emplacement de la table de facturation se base sur le paragraphe [Table] présent dans le template word
        /// </summary>
        /// <param name="docGen">Le <see cref="WordprocessingDocument"/> utilisé pour piloter le document word</param>
        /// <returns>True si tout c'est bien passé</returns>
        private bool InsertDataIntoTable(WordprocessingDocument docGen)
        {
            try
            {
                Table lignesfactures = GenerateTableAndHeader(docGen);
                float TotHtva = 0;
                float Tottva = 0;
                if (lignesfactures != null)
                {

                    foreach (IDataFactures item in _ldatas)
                    {
                        List<OpenXmlElement> listcells = new List<OpenXmlElement>();
                        foreach (PropertyInfo prop in typeof(IDataFactures).GetProperties()
                       .OrderBy(p => (p.GetCustomAttribute<JsonPropertyOrderAttribute>()?.Order)))
                        {
                            string? val = prop.GetValue(item)?.ToString();
                            listcells.Add(new TableCell(new Paragraph(new Run(new Text(val ?? "ERROR")))));

                        }
                        TableRow tr1 = new TableRow(listcells);

                        // Add row to the table.
                        lignesfactures.AppendChild(tr1);
                        TotHtva += (item.PRIX_UNITAIRE_HT * item.QUANTITE);
                        Tottva += (TotHtva + (TotHtva * (((float)item.TAUX_TVA) / 100)));

                    }
                    int nbCells = typeof(IDataFactures).GetProperties().Count();
                    List<OpenXmlElement> cells = new List<OpenXmlElement>();
                    for (int i = 0; i < nbCells - 3; i++)
                    {
                        cells.Add(new TableCell(new Paragraph(new Run(new Text("")))));
                    }
                    cells.Add(new TableCell(new Paragraph(new Run(new Text("Total")))));
                    cells.Add(new TableCell(new Paragraph(new Run(new Text($"{TotHtva} € (HTVA)")))));
                    cells.Add(new TableCell(new Paragraph(new Run(new Text($"{Tottva} € (TTC)")))));
                    TableRow tr2 = new TableRow(cells);

                    // Add row to the table.

                    lignesfactures.AppendChild(tr2);


                }


                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Permet de générer la table avec le header. Le header de la table sera en couleur.
        /// Le contenu des cellules se basent sur l'attribut <see cref="JsonPropertyNameAttribute"/> de l'interface <see cref="IDataFactures"/>
        /// et l'ordre des cellules se basent sur l'attribut <see cref="JsonPropertyOrderAttribute"/> de l'interface <see cref="IDataFactures"/>
        /// </summary>
        /// <param name="docGen">Le <see cref="WordprocessingDocument"/> utilisé pour piloter le document word</param>
        /// <returns>La référence vers la table insérée dans le document word</returns>
        private Table GenerateTableAndHeader(WordprocessingDocument docGen)
        {   
            //Récupérer l'emplacement du Text [Table]
            Body body = docGen.MainDocumentPart.Document.Body;
            var t = body.Descendants<Paragraph>().First<Paragraph>(p => p.InnerText.Equals("[Table]"));
            if(t!=null)
            {
                
               Table tbl = new Table();
                // Set the style and width for the table.
               TableProperties tblProp = new TableProperties
                    (
                       new TableBorders(
                                           new TopBorder()
                                           {
                                               Val =
                                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                                               Size = 10
                                           },
                                           new BottomBorder()
                                           {
                                               Val =
                                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                                               Size = 10
                                           },
                                           new LeftBorder()
                                           {
                                               Val =
                                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                                               Size = 10
                                           },
                                           new RightBorder()
                                           {
                                               Val =
                                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                                               Size = 10
                                           },
                                           new InsideHorizontalBorder()
                                           {
                                               Val =
                                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                                               Size = 10
                                           },
                                           new InsideVerticalBorder()
                           {
                               Val =
                               new EnumValue<BorderValues>(BorderValues.BasicThinLines),
                               Size = 10
                           }
                                        )                
                        );              
                
                TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
                // Make the table width 100% of the page width.
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
                // Apply
                tblProp.Append(tableStyle, tableWidth);
                tbl.AppendChild(tblProp);

                
                List<OpenXmlElement> lcells = new List<OpenXmlElement>();
                foreach (PropertyInfo item in typeof(IDataFactures).GetProperties()
                    .OrderBy(p=>(p.GetCustomAttribute<JsonPropertyOrderAttribute>()?.Order)))
                {
                    
                    TableCell cell = new TableCell(new Paragraph(new Run(new Text(item.GetCustomAttribute<JsonPropertyNameAttribute>()?.Name ?? item.Name))));
                    TableCellProperties style = new TableCellProperties();                 
                    var shading = new Shading()
                    {
                        Color = "auto",
                        Fill = "ABCDEF",
                        Val = ShadingPatternValues.Percent25
                    };
                    style.Append(shading);
                    cell.Append(style);
                    lcells.Add(cell);
                } 
                 
                 
                //Add header
                TableRow tr1 = new TableRow(lcells);

                 
               

                tbl.AppendChild(tr1);
                t.InsertAfterSelf(tbl);
                t.Remove();
                return tbl;
            }

            return new Table();
        }

        /// <summary>
        /// Permet de récupérer les valeurs pour le fiel basé sur son nom dans Waord
        /// </summary>
        /// <param name="FieldName">Le nom du field dans word</param>
        /// <returns>Le string récupéré</returns>
        /// <exception cref="Exception">Si la valeur du FieldName ne se trouve pas dans le dictionnaire transmis dans le constructeur</exception>
        private string GetMergeValue(string FieldName)
        {
            if (_fieldValues.ContainsKey(FieldName))
            {
                return _fieldValues[FieldName];
            }
                else
            { 
                throw new Exception(message: "FieldName (" + FieldName + ") was not found");
            }
        }

        /// <summary>
        /// Permet de remplacer le field de publipostage par la valeur transmise
        /// </summary>
        /// <param name="field">le field de publipostage</param>
        /// <param name="replacementText">Le texte a inserer</param>
        private void ReplaceMergeFieldWithText(FieldCode field, string replacementText)
        {
            if (field == null || replacementText == string.Empty)
            {
                return;
            }           

            Run rFldParent = (Run)field.Parent;
            List<Run> runs = new List<Run>();

            runs.Add(rFldParent.PreviousSibling<Run>());  
            runs.Add(rFldParent.NextSibling<Run>());  

            foreach (Run run in runs)
            {
                run.Remove();
            }

            field.Remove();  
            rFldParent.Append(new Text(replacementText));
        }




    }
}