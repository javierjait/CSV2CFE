using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using System.Configuration;
using Microsoft.VisualBasic.FileIO;
using Newtonsoft.Json;
using System.Globalization;
using System.ServiceModel.Configuration;

namespace CSV2CFE
{
    class JSONValue
    {
        public string Value
        {
            get; set;
        }
    }

    class FactObj // Objeto factura
    {
        public string ID { get; set; }
        public string Cliente { get; set; }
        public string IDFisc { get; set; }
        public string Moneda { get; set; }
        public Decimal Monto { get; set; }
        public string Descripcion { get; set; }
        public string Domicilio { get; set; }
        public string Localidad { get; set; }
        public string Departamento { get; set; }
        public string Pais { get; set; }
        public string CodPais { get; set; }
        public string CP { get; set; }
        public decimal Cotizacion { get; set; }
    }

    class Utf8StringWriter : StringWriter // Utf8 writer
    {
        public override Encoding Encoding { get { return Encoding.UTF8; } }
    }

    class Program
    {
        static private string _nomUsuario;
        static private string _clave;
        static private string _tenant;
        static private string _adenda;
        static private readonly string _pathEnviados = "Enviados";
        static private string _pathFacturas = "Facturas_PDF\\";
        //static private readonly decimal _ui = 4.7846M; // Unidad indexada

        static void Main(string[] args)
        {
            ReadConfigurarion();

            if (string.IsNullOrWhiteSpace(_nomUsuario) || string.IsNullOrWhiteSpace(_clave) || string.IsNullOrWhiteSpace(_tenant))
            {
                AddToLog("Los valores de configuración con incorrectos");
                return;
            }

            if (CheckMiddleware())
            {
                SearchCSV();
                /// Test check
                /*string endpointAddress = GetEndpointAddress();
                if (endpointAddress == "http://testing.cloud.sicfe.uy:8081/sicfe/Services/sicfe.svc")
                {
                    SearchCSV();
                }
                else
                {
                    AddToLog("El endpoint address es incorrecto");
                }*/
            }
        }

        /// <summary>
        /// Lee la configuración del .config
        /// </summary>
        static private void ReadConfigurarion()
        {
            _nomUsuario = ConfigurationManager.AppSettings["usuario"];
            _clave = ConfigurationManager.AppSettings["clave"];
            _tenant = ConfigurationManager.AppSettings["tenant"];
            _adenda = ConfigurationManager.AppSettings["adenda"];
            string configFolder = ConfigurationManager.AppSettings["carpetaPDF"];
            if (configFolder != null) // Agrega la barra al final de la carpeta si no existe
            {
                // Agrega la barra final a la carpeta si es necesario
                if (configFolder.EndsWith("\\"))
                    _pathFacturas = configFolder;
                else
                    _pathFacturas = configFolder + "\\";
            }

            if (!Directory.Exists(_pathFacturas)) // Crea carpeta si no existe
            {
                Directory.CreateDirectory(_pathFacturas);
            }
        }

        /// <summary>
        /// Verifica la conexión al middleware
        /// </summary>
        /// <returns></returns>
        static private bool CheckMiddleware()
        {
            srSiCFE.SICFEEmisorClient ws = new srSiCFE.SICFEEmisorClient();
            try
            {
                var res = ws.PingSICFEEmisor(_nomUsuario, _clave, _tenant);
                if (res.Codigo != 0)
                {
                    AddToLog(res.Descripcion);
                }

                return res.Codigo == 0;
            }
            catch (Exception ex)
            {
                AddToLog(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Busca los archivos CSV
        /// </summary>
        static private void SearchCSV()
        {
            string exeFolder = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName); // Obtiene la carpeta donde se ejecuta la aplicación
            string rucEmisor = ConfigurationManager.AppSettings["RUCEmisor"];   // Obtiene el RUCEmisor
            srSiCFE.SICFEEmisorClient ws = new srSiCFE.SICFEEmisorClient(); // Inicializa el ws            

            if (!string.IsNullOrEmpty(rucEmisor)) // Si existe RUCEmisor, continua
            {
                var archivos = Directory.GetFiles(exeFolder, "*.csv"); // Obtiene los archivos CSV en la carpeta
                foreach (string archivo in archivos) // Iteractua los archivos encontrados
                {
                    //Console.WriteLine(archivo);
                    var flist = getCSV(archivo); // Genera la lista de facturas

                    foreach (var f in flist) // Iteractua las facturas
                    {
                        string xml = generateXML(f);

                        if (!string.IsNullOrWhiteSpace(xml)) // Verifica que XML no esté vacío
                        {
                            srSiCFE.SICFERespuestaEnvioCFE res;

                            string refERP = $"{f.ID}";

                            try
                            {
                                res = ws.EnvioCFE(_nomUsuario, _clave, _tenant, "", xml, refERP, "", false, 22, 0, "", "", true, false, ""); // Llama al middleware
                                if (res.Codigo == 0 || res.Codigo == 100009) // Factura generada OK, o ya existe el refERP
                                {
                                    var pdf = ws.ObtenerPDF(_nomUsuario, _clave, res.IdCFE, _tenant, "");
                                    File.WriteAllBytes($"{_pathFacturas}{refERP}.pdf", pdf.Buffer);
                                    if (res.Codigo == 100009) // Existe el refERP, agrega al log
                                    {
                                        AddToLog($"ID: {refERP}, {res.Descripcion}");
                                    }
                                }
                                else if (res.Codigo == 202) // Error de credenciales
                                {
                                    AddToLog($"{res.Descripcion}");
                                    return; // Finaliza
                                }
                                else // Otro error
                                {
                                    AddToLog($"ID: {refERP}, {res.Descripcion}");
                                }
                            }
                            catch (Exception e)
                            {
                                AddToLog(e.Message);
                            }
                        }
                    }

                    try
                    {
                        Directory.CreateDirectory(_pathEnviados); // Crea carpeta
                        File.Move(archivo, $"{_pathEnviados}\\{DateTime.Now:yyyyMMdd_HHmm}_{Path.GetFileName(archivo)}"); // Mueve archivo CSV
                    }
                    catch (Exception e)
                    {
                        AddToLog(e.Message);
                    }
                }
            }
        }

        /// <summary>
        /// Genera el objeto de factura
        /// </summary>
        /// <param name="archivo">archivo CSV</param>
        /// <returns>lista de objeto factura</returns>
        static private List<FactObj> getCSV(string archivo)
        {
            List<FactObj> fact = new List<FactObj>();

            using (TextFieldParser csvParser = new TextFieldParser(archivo)) // Crea el parserCSV
            {
                csvParser.SetDelimiters(new string[] { "," }); // Genera el delimitador
                csvParser.HasFieldsEnclosedInQuotes = true;

                var header = csvParser.ReadFields(); // Lee el encabezado                 
                var headerDictionary = createFieldsDictionary(header); // Genera el diccionario del encabezado

                while (!csvParser.EndOfData) // Lee las facturas hasta el final
                {
                    string[] fields = csvParser.ReadFields(); // Obtiene los campos
                    string id = "", descripcion = "", IDFiscal = "";
                    JSONValue cliente = new JSONValue(), moneda = new JSONValue(), domicilio = new JSONValue(), localidad = new JSONValue(),
                              departamento = new JSONValue(), cp = new JSONValue(), pais = new JSONValue(), codPais = new JSONValue();
                    decimal monto = 0, cotizacion = 1;

                    foreach (KeyValuePair<string, int> column in headerDictionary) // Busca los campos y los asigna a la factura
                    {
                        switch (column.Key)
                        {
                            case "ID":
                                id = fields[column.Value];
                                break;
                            case "CLIENTE":
                                cliente = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "CONSUMO": // ID Fiscal
                                JSONValue IDFisc = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                if (!string.IsNullOrEmpty(IDFisc.Value))
                                {
                                    IDFiscal = IDFisc.Value;
                                }
                                break;
                            case "MONEDA":
                                moneda = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "DOMICILIO":
                                domicilio = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "LOCALIDAD":
                                localidad = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "DEPARTAMENTO":
                                departamento = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "CP":
                                cp = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "PAIS":
                                codPais = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "NPAIS":
                                pais = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                break;
                            case "MONTO":
                                decimal.TryParse(fields[column.Value].Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out monto);
                                monto = Math.Round(monto, 2);
                                break;
                            case "COTIZACION":
                                JSONValue cotizTmp = JsonConvert.DeserializeObject<JSONValue>(fields[column.Value]);
                                if (cotizTmp != null)
                                {
                                    decimal.TryParse(cotizTmp.Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out cotizacion);
                                }
                                break;
                            case "DESCRIPCION":
                                /*Regex regex = new Regex("<P>(.*)</P>", RegexOptions.IgnoreCase);
                                var v = regex.Match(fields[column.Value]);
                                if (v.Groups.Count > 1)
                                {
                                    descripcion = StripHTML(v.Groups[1].ToString());
                                }*/
                                descripcion = fields[column.Value];
                                break;
                        }
                    }

                    if (!string.IsNullOrEmpty(id)) // Verifica que exista el ID
                    {
                        fact.Add(new FactObj
                        {
                            ID = id,
                            Cliente = cliente.Value,
                            IDFisc = IDFiscal,
                            Moneda = moneda.Value,
                            Monto = monto,
                            Descripcion = descripcion,
                            Domicilio = domicilio.Value,
                            Localidad = localidad.Value,
                            Departamento = departamento.Value,
                            CP = cp.Value,
                            Pais = pais.Value,
                            CodPais = codPais.Value,
                            Cotizacion = cotizacion
                        }); // Agrega la factura a la lista
                    }
                }
            }

            return fact;
        }


        /// <summary>
        /// Busca los campos y obtiene el número de ubicación
        /// </summary>
        /// <param name="header">Array del header</param>
        /// <returns></returns>
        static private Dictionary<string, int> createFieldsDictionary(string[] header)
        {
            Dictionary<string, int> fieldsDictionary = new Dictionary<string, int>(); // Crea el diccionario
            for (int i = 0; i < header.Length; i++) // Iteractua los elementos del header
            {
                if (header[i] == "ID" || header[i] == "CLIENTE" || header[i] == "MONEDA" || header[i] == "MONTO")
                {
                    fieldsDictionary.Add(header[i], i); // Agrega el campo y la posición
                }
                else if (header[i] == "MONEDAYCOTIZACION_x003a_COTIZACI")
                {
                    fieldsDictionary.Add("COTIZACION", i); // Agrega el campo y la posición
                }
                else if (header[i] == "Texto_x0020_Linea_x0020_Factura")
                {
                    fieldsDictionary.Add("DESCRIPCION", i); // Agrega el campo y la posición
                }
                /*else if (header[i].StartsWith("DESCRIPCION"))
                {
                    fieldsDictionary.Add("DESCRIPCION", i);
                }*/
                else
                {
                    bool startCliente = header[i].StartsWith("CLIENTE");
                    bool endId = !header[i].EndsWith("#Id");
                    string campo = checkCampos(header[i]); // Verifica el campo
                    if (startCliente && endId && !string.IsNullOrEmpty(campo))
                    {
                        fieldsDictionary.Add(campo, i); // Agrega el campo y la posición
                    }
                }
            }
            return fieldsDictionary;
        }


        /// <summary>
        /// Verifica que el campo esté en el listado
        /// </summary>
        /// <param name="field">Campo a buscar</param>
        /// <returns></returns>
        static private string checkCampos(string field)
        {
            string[] campos = { "DOMICILIO", "LOCALIDAD", "DEPARTAMENTO", "CP", "NPAIS", "PAIS", "CONSUMO" };

            foreach (string campo in campos)
            {
                if (field.Contains(campo))
                {
                    return campo;
                }
            }

            return null;
        }


        static private string StripHTML(string input)
        {
            string item = Regex.Replace(input, "<.*?>", String.Empty);
            item = Regex.Replace(item, @"[^\p{L}\p{Nd}\s-_#/]", String.Empty);

            return item.Trim();
        }

        /// <summary>
        /// General el XML
        /// </summary>
        /// <param name="factura">Objeto factura</param>
        /// <returns></returns>
        private static string generateXML(FactObj factura)
        {
            CFE_Empresas_Type cfe = new CFE_Empresas_Type();
            CFEDefType cfeDT = new CFEDefType();
            CFEDefTypeETck cfeTck = new CFEDefTypeETck();
            CFEDefTypeETckEncabezado encabezado = new CFEDefTypeETckEncabezado();
            IdDoc_Tck idDoc = new IdDoc_Tck();
            Item_Det_Fact det = new Item_Det_Fact();
            CAEDataType CAEdata = new CAEDataType();
            Emisor emisor = new Emisor();
            Receptor_Tck receptor = new Receptor_Tck();
            Totales totales = new Totales();

            string Xml = "";

            try
            {

                idDoc.TipoCFE = IdDoc_TckTipoCFE.Item101;
                idDoc.Nro = "1";
                idDoc.Serie = "1";
                idDoc.FchEmis = DateTime.Today; // Fecha de hoy

                emisor.RUCEmisor = ConfigurationManager.AppSettings["RUCEmisor"]; // settings de la aplicación
                emisor.RznSoc = ConfigurationManager.AppSettings["RznSoc"]; // settings de la aplicación

                string CdgDGISucur = ConfigurationManager.AppSettings["CdgDGISucur"]; // settings de la aplicación
                if (string.IsNullOrEmpty(CdgDGISucur))
                {
                    emisor.CdgDGISucur = "3";
                }
                else
                {
                    emisor.CdgDGISucur = CdgDGISucur;
                }

                emisor.DomFiscal = ConfigurationManager.AppSettings["DomFiscal"]; // settings de la aplicación
                emisor.Ciudad = ConfigurationManager.AppSettings["Ciudad"];       // settings de la aplicación
                emisor.Departamento = ConfigurationManager.AppSettings["Departamento"]; // settings de la aplicación
                emisor.InfoAdicionalEmisor = ConfigurationManager.AppSettings["InfoAdicional"]; // settings de la aplicación

                receptor.RznSocRecep = factura.Cliente;
                receptor.CiudadRecep = factura.Localidad;
                receptor.PaisRecep = factura.Pais;

                //receptor.CP = factura.CP;
                try
                {
                    receptor.CP = Regex.Replace(factura.CP, @"\D", "");
                }
                catch
                {
                    AddToLog("CP: Valor inválido o nulo");
                    return "";
                }
                receptor.DeptoRecep = factura.Departamento;
                receptor.DirRecep = factura.Domicilio;
                receptor.PaisRecep = factura.Pais;
                if (!string.IsNullOrWhiteSpace(factura.IDFisc)) // Verifica el ID fiscal
                {
                    try
                    {
                        receptor.CodPaisRecep = (CodPaisType)Enum.Parse(typeof(CodPaisType), factura.CodPais.ToUpper().Substring(0, 2));
                        receptor.CodPaisRecepSpecified = true;
                    }
                    catch (Exception e)
                    {
                        AddToLog(e.Message);
                        return ""; // Si el código de país no es válido devuelve XML en blanco
                    }
                    if (factura.CodPais == "UY")
                    {
                        receptor.ItemElementName = ItemChoiceType.DocRecep;
                        string doc = Regex.Replace(factura.IDFisc, @"[^\d]", ""); // solo números
                        if (isRutValid(doc)) // Verifica el RUC
                        {
                            receptor.TipoDocRecep = DocType.Item2; // RUC
                            receptor.Item = doc.Substring(0, Math.Min(12, doc.Length));
                        }
                        else
                        {
                            receptor.TipoDocRecep = DocType.Item3; // Cedula
                            receptor.Item = doc.Substring(0, Math.Min(20, doc.Length));
                        }
                    }
                    else
                    {
                        receptor.TipoDocRecep = DocType.Item7;
                        receptor.ItemElementName = ItemChoiceType.DocRecepExt;
                        receptor.Item = factura.IDFisc.Substring(0, Math.Min(20, factura.IDFisc.Length)); // hast 20 caracteres
                    }
                    receptor.TipoDocRecepSpecified = true;
                }

                totales.CantLinDet = "1";
                try
                {
                    totales.TpoMoneda = (TipMonType)Enum.Parse(typeof(TipMonType), factura.Moneda); //TipMonType.UYU;
                }
                catch (Exception e)
                {
                    AddToLog(e.Message);
                }
                totales.MntNoGrv = factura.Monto;
                totales.MntNoGrvSpecified = true;
                totales.MntTotal = factura.Monto;
                totales.MntPagar = factura.Monto;
                if (totales.TpoMoneda != TipMonType.UYU)
                {
                    totales.TpoCambio = Math.Round(factura.Cotizacion, 3); // Ver tipo de cambio
                    totales.TpoCambioSpecified = true;
                }

                encabezado.IdDoc = idDoc;
                encabezado.Emisor = emisor;
                encabezado.Receptor = receptor;
                encabezado.Totales = totales;

                det.NroLinDet = "1";
                det.MontoItem = factura.Monto;
                det.NomItem = factura.Descripcion.Substring(0, Math.Min(80, factura.Descripcion.Length));
                det.PrecioUnitario = factura.Monto;
                det.UniMed = "N/A";
                det.Cantidad = 1;

                List<Item_Det_Fact> detalle = new List<Item_Det_Fact>();
                detalle.Add(det);

                CAEdata.CAE_ID = "";
                CAEdata.DNro = "";
                CAEdata.HNro = "";
                /*CAEdata.CAEEspecial = "0";
                CAEdata.CAEEspecialSpecified = true;
                CAEdata.CausalCAEEsp = "0";*/
                CAEdata.FecVenc = "";

                cfeTck.Encabezado = encabezado;
                cfeTck.Detalle = detalle.ToArray();
                cfeTck.CAEData = CAEdata;
                cfeDT.Item = cfeTck;
                cfe.CFE = cfeDT;

                if (string.IsNullOrEmpty(_adenda))
                {
                    cfe.Adenda = new Object();
                }
                else
                {
                    cfe.Adenda = _adenda;
                }

                XmlSerializer serializer = new XmlSerializer(typeof(CFE_Empresas_Type));

                using (var st = new Utf8StringWriter())
                {
                    serializer.Serialize(st, cfe); // Serializa el XML
                    Xml = st.ToString();           // XML como string
                }
            }
            catch (Exception ex)
            {
                AddToLog(ex.Message);
            }

            return Xml;
        }

        /// <summary>
        /// Verifica si el RUT es válido
        /// </summary>
        /// <param name="rut">RUT</param>
        /// <returns></returns>
        static private bool isRutValid(string rut)
        {
            if (rut.Length != 12)
            {
                return false;
            }

            Regex reg = new Regex("^[0-9]+$");

            if (!reg.Match(rut).Success)
            {
                return false;
            }

            int dc = Int32.Parse(rut.Substring(11));
            //string rutAux = rut.Substring(0, 11);
            int total = 0;
            int factor = 2;

            for (int i = 10; i >= 0; i--)
            {
                total += (factor * Int32.Parse(rut.Substring(i, 1)));
                factor = factor == 9 ? 2 : ++factor;
            }

            int dv = 11 - (total % 11);
            if (dv == 11)
            {
                dv = 0;
            }
            else if (dv == 10)
            {
                dv = 1;
            }

            return dv == dc;
        }

        /// <summary>
        /// Agrega el mensaje al archivo log.txt
        /// </summary>
        /// <param name="msg">Mensaje a escribir</param>
        static private void AddToLog(string msg)
        {
            try
            {
                File.AppendAllText("log.txt", $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}: {msg}\n"); // Escribe fecha/hora y mensaje
            }
            catch 
            {
                //Console.WriteLine(ex.Message);
            }
        }

        static private string GetEndpointAddress()
        {
            ClientSection clientSect = (ClientSection)ConfigurationManager.GetSection("system.serviceModel/client");
            if (clientSect == null)
            {
                return string.Empty;
            }
            ChannelEndpointElementCollection endpoints = (ChannelEndpointElementCollection)clientSect.ElementInformation.Properties[string.Empty].Value;
            if (endpoints.Count == 0)
            {
                return string.Empty;
            }

            return endpoints[0].Address.ToString();
        }
    }
}
