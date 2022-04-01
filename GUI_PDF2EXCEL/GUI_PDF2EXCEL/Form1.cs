using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

using System.IO;
using System.Diagnostics;
using iTextSharp.text.pdf;
using System.Globalization;

using _Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Collections;
using System.Net.Mail;
using System.Net.Mime;
using System.Xml;

namespace GUI_PDF2EXCEL
{
    public partial class Form1 : Form
    {
        public static string Proyecto = null;
        public static string Cliente = null;
        public static string Tecnico = null;
        public static string Semana = null;
        public static string SemanaCSV = null;
        public static string Estado=null;
        public static string TipoIntervencion = null;
        public static string AccionesRealizadas=null;
        public static string Observaciones = null;
        public static string MotivoIntervencion = null;
        public static int month;
        public static string[] Dia = new string[7];
        public static int[] Fechas = new int[7];
        public static string[] FechaCompleta = new string[7];
        public static int[] Mes = new int[7];
        //--------------------------------------------
        public static string[] Semanas = new string[7];
        public static string[] Clientes = new string[7];
        public static string[] Proyectos = new string[7];
        public static double[] dHNormal = new double[7];
        public static double[] dHExtra = new double[7];
        public static double[] dHFestiva = new double[7];
        public static double[] dHZombie = new double[7];
        public static double[] dHViaje = new double[7];
        public static double[] dHSumas = new double[7];
        //--------------------------------------------
        public static string[] hNormal = new string[7];
        public static string[] hExtra = new string[7];
        public static string[] hFestiva = new string[7];
        public static string[] hZombie = new string[7];
        public static string[] hViaje = new string[7];
        public static string[] totales = new string[7];
        //---------------------------------------------
        public static string[] hsp = new string[7];
        public static string[] hlp = new string[7];
        public static string[] hip = new string[7];
        public static string[] hfp = new string[7];
        public static string[] his = new string[7];
        public static string[] hfs = new string[7];
        public static string[] hss = new string[7];
        public static string[] hls = new string[7];
        //---------------------------------------------
        public static string pdfTemplate;
        public static string csvTemplate;
        public static PdfReader pdfReader;
        public static AcroFields form;
        public static string pathExcel;
        public static string filePathCSV;
        public static string nameFileCSV;
        public static _Excel.Application xls;
        public static _Excel.Workbook wkbk;
        public static _Excel.Sheets sheet;
        public static string currentSheet;
        public static _Excel.Worksheet excelWorksheet;
        public static _Excel.Range excelCell;
        public static string usp;
        //--------------------------------------------
        private static XmlDocument doc1;
        public static string lastDirPDF;
        public static string lastDirExcel;
        public static string lastDirCSV;
        public static string pathXml;
        //--------------------------------------------
        public  const double Dia_Normal=8.5;
        public const double Horas_Zombie = 8;
        public  const double Dia_Viernes = 5;
        private object Globals;

        //--------------------------------------------

        public Form1()
        {
            InitializeComponent();
            pathExcel = "NOK";
            pdfTemplate = "NOK";
            String dirXml = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
            String fileXml = "resources\\user_data.xml";
            pathXml = Path.Combine(dirXml, fileXml);
            doc1 = new XmlDocument();
            doc1.Load(pathXml);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.getPathExcel();
            textBox2.Text = pathExcel;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.getPathPDF();
            textBox1.Text = pdfTemplate;
        }
        public void export()
        {

            if (listBox1.SelectedIndex != -1)
            {

                month = this.getMonth() + 1;


                if ((pathExcel != "NOK") && (pdfTemplate != "NOK") && (textBox1.Text != "") && (textBox2.Text != ""))
                {
                    try
                    {

                        this.getFormPdf();
                        this.getFields();
                        this.getExcelCell();
                        this.prepareData();
                        this.setCells();

                        for (int i = 0; i < 7; i++)
                        {
                            Fechas[i] = 0;
                        }


                        // Save the changes and close the workbook.
                        pdfReader.Close();
                        wkbk.Close(true, Type.Missing, Type.Missing);
                        xls.Quit();

                        MessageBox.Show(" -- Proceso Finalizado! -- ");

                    }
                    catch (System.NullReferenceException e)
                    {

                        MessageBox.Show(" Error: Proporcionar una ruta correcta para el pdf \n\n" + e);


                    }
                    catch (System.Runtime.InteropServices.COMException e)
                    {

                        MessageBox.Show(" Error: Proporcionar una ruta correcta para el excel \n\n" + e);


                    }
                    catch (System.Exception e)
                    {
                        pdfReader.Close();
                        wkbk.Close(true, Type.Missing, Type.Missing);
                        xls.Quit();
                        MessageBox.Show(" Error: Se ha producido un error grave \n\n" + e);


                    }


                }
                else
                {
                    MessageBox.Show(" Error: Ruta de archivos incorrecta.\n Utilice los botones de seleccion de ruta de archivos:\n * Origen PDF\n * Destino Excel ");
                }
            }
            else
            {
                MessageBox.Show(" Error: Seleccione mes para el que se van a extraer datos.\n Utlice la lista de la derecha par seleccionar mes ");
            }            

        }
        static bool IsOpened(string aPath)
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream(aPath, FileMode.Open, FileAccess.Read, FileShare.None);
                return false;
            }
            catch (IOException )
            {
                return true;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }
        }
        public void exportCSV()
        {
            //string tempHoraInit=null;
            //string tempHoraFin=null;
            if(csvTemplate != "NOK")
            {
                try
                {
                    this.getFormPdf();
                    this.getFields4CSV();

                    Regex initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z]* ?");
                    string init = initials.Replace(Tecnico, "$1");
                    nameFileCSV = init + Proyecto + "_W"+ SemanaCSV;
                    filePathCSV = csvTemplate + "\\" + nameFileCSV.Trim() + ".csv";
                    StringBuilder sb = new StringBuilder();

                    ArrayList DatosProyecto = new ArrayList();
                    ArrayList resumenHoras = new ArrayList();
                    ArrayList tramasHorarias = new ArrayList();

                    ArrayList cmDatosProyecto = new ArrayList();
                    ArrayList cmResumenHoras = new ArrayList();
                    ArrayList cmTramasHorarias = new ArrayList();

                    //----------------------------------------------------
                    //-----Comentarios Datos proyecto---------------------
                    string[] cmCamposDatosProyecto = new string[10];
                    //----------------------------------------
                    cmCamposDatosProyecto[0] = "#Fecha Creacion";
                    cmCamposDatosProyecto[1] = "#Codigo Trabajador";
                    cmCamposDatosProyecto[2] = "#Nombre Tecnico";
                    cmCamposDatosProyecto[3] = "#Proyecto";
                    cmCamposDatosProyecto[4] = "#Semana";
                    cmCamposDatosProyecto[5] = "#Estado";
                    cmCamposDatosProyecto[6] = "#TipoIntervencion";
                    cmCamposDatosProyecto[7] = "#MotivoIntervencion";
                    cmCamposDatosProyecto[8] = "#AccionesRealizadas";
                    cmCamposDatosProyecto[9] = "#Observaciones";
                    cmDatosProyecto.Add(string.Join(";", cmCamposDatosProyecto));
                    //----------------------------------------------------
                    //Comentarios Tramas horarias write csv
                    //----------------------------------------------------
                    string[] cmCamposTramasHorarias = new string[9];
                    //----------------------------------------------------
                    cmCamposTramasHorarias[0] = "#Fecha";
                    cmCamposTramasHorarias[1] = "#Hora Salida 1";
                    cmCamposTramasHorarias[2] = "#Hora Llegada 1";
                    cmCamposTramasHorarias[3] = "#Hora Inicio 1";
                    cmCamposTramasHorarias[4] = "#Hora Fin 1";
                    cmCamposTramasHorarias[5] = "#Hora Inicio 2";
                    cmCamposTramasHorarias[6] = "#Hora Fin 2";
                    cmCamposTramasHorarias[7] = "#Hora Salida 2";
                    cmCamposTramasHorarias[8] = "#Hora Llegada 2";
                    cmTramasHorarias.Add(string.Join(";", cmCamposTramasHorarias));
                    //----------------------------------------------------
                    //Comentarios Resumen horarias write csv
                    //----------------------------------------------------
                    string[] cmCamposResumenHoras = new string[7];
                    //----------------------------------------
                    cmCamposResumenHoras[0] = "#Fecha";
                    cmCamposResumenHoras[1] = "#Hora N";
                    cmCamposResumenHoras[2] = "#Hora E";
                    cmCamposResumenHoras[3] = "#Hora F";
                    cmCamposResumenHoras[4] = "#Hora Z";
                    cmCamposResumenHoras[5] = "#Hora V";
                    cmCamposResumenHoras[6] = "#Total Horas";
                    cmResumenHoras.Add(string.Join(";", cmCamposResumenHoras));
                    //----------------------------------------------------                   
                    //Datos proyecto
                    //----------------------------------------------------
                    string[] camposDatosProyecto = new string[10];
                    //----------------------------------------
                    camposDatosProyecto[0] = DateTime.Now.ToString("dd/MM/yyyy");
                    camposDatosProyecto[1] = getCodOper();
                    camposDatosProyecto[2] = Tecnico.Replace(";","");
                    camposDatosProyecto[3] = Proyecto;
                    camposDatosProyecto[4] = SemanaCSV;
                    camposDatosProyecto[5] = Estado;
                    camposDatosProyecto[6] = TipoIntervencion.Remove(TipoIntervencion.Length-1);
                    camposDatosProyecto[7] = MotivoIntervencion.Replace(";", ":");
                    camposDatosProyecto[8] = AccionesRealizadas.Replace(";", ":");
                    camposDatosProyecto[9] = Observaciones.Replace(";", ":");
                    DatosProyecto.Add(string.Join(";", camposDatosProyecto));                                       
                    //----------------------------------------                    
                    //Tramas horarias write csv
                    //----------------------------------------
                    for (int i = 0; i < 7; i++)
                    {
                        string[] campos = new string[9];
                        if (FechaCompleta[i] != "")
                        {
                            //---------------------
                            campos[0] = FechaCompleta[i];
                            campos[1] = hsp[i];
                            campos[2] = hlp[i];
                            campos[3] = hip[i];
                            campos[4] = hfp[i];
                            campos[5] = his[i];
                            campos[6] = hfs[i];
                            campos[7] = hss[i];
                            campos[8] = hls[i];
                            tramasHorarias.Add(string.Join(";", campos));
                        }
                    }
                    //----------------------------------------
                    //Resumen horas write csv
                    //----------------------------------------
                    for (int i = 0; i < 7; i++)
                    {
                        string[] campos = new string[7];
                        if (FechaCompleta[i] != "")
                        {
                            //---------------------
                            campos[0] = FechaCompleta[i];
                            campos[1] = hNormal[i];
                            campos[2] = hExtra[i];
                            campos[3] = hFestiva[i];
                            campos[4] = hZombie[i];
                            campos[5] = hViaje[i];
                            campos[6] = totales[i];                            
                            resumenHoras.Add(string.Join(";", campos));
                        }
                    }
                    //############################################
                    //------Comentarios Datos proyecto ----------
                    int length01 = cmDatosProyecto.Count;
                    for (int index = 0; index < length01; index++)
                    {
                        sb.AppendLine(cmDatosProyecto[index].ToString());
                    }
                    //Datos proyecto                   
                    int length0 = DatosProyecto.Count;
                    for (int index = 0; index < length0; index++)
                    {
                        sb.AppendLine(DatosProyecto[index].ToString());
                    }
                    //------Comentarios Tramas horarias ----------
                    int length02 = cmTramasHorarias.Count;
                    for (int index = 0; index < length02; index++)
                    {
                        sb.AppendLine(cmTramasHorarias[index].ToString());
                    }
                    //Tramas horarias write csv
                    int length1 = resumenHoras.Count;
                    for (int index = 0; index < length1; index++)
                    {
                        sb.AppendLine(tramasHorarias[index].ToString());
                    }
                    //------Comentarios Resumen horas ----------
                    int length03 = cmResumenHoras.Count;
                    for (int index = 0; index < length03; index++)
                    {
                        sb.AppendLine(cmResumenHoras[index].ToString());
                    }
                    //Resumen horas write csv
                    int length2 = resumenHoras.Count;
                    for (int index = 0; index < length2; index++)
                    {
                        sb.AppendLine(resumenHoras[index].ToString());
                    }


                    File.WriteAllText(filePathCSV, CalcMD5Hash(sb).ToString().TrimEnd('\r', '\n'));
                    for (int i = 0; i < 7; i++)
                    {
                        FechaCompleta[i] = "";
                    }

                    pdfReader.Close();

                    MessageBox.Show(" -- Proceso finalizado! -- ");

                    DialogResult result = MessageBox.Show("Se ha generado el archivo .csv\n ¿Quiere enviar el archivo por email?", "Atencion",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.Yes)
                    {
                        this.sendAttachment();
                    }
                    else if (result == DialogResult.No)
                    {
                        //code for No
                    }           
                }
                catch (System.NullReferenceException e)
                {

                    MessageBox.Show(" Error: Proporcionar una ruta correcta para el pdf \n\n" + e);


                }                
                catch (System.Exception e)
                {
                    pdfReader.Close();
                    MessageBox.Show(" Error: Se ha producido un error grave \n\n" + e);
                }


            }
            else
            {
               
                {
                   MessageBox.Show(" Error: Seleccione una ruta correcta para archivo .csv  ");
            
                }
                
            }

        }

        public static StringBuilder CalcMD5Hash(StringBuilder aSb)
        {
            StringBuilder sb = new StringBuilder();

            System.Security.Cryptography.MD5 md5 = System.Security.Cryptography.MD5.Create();

            byte[] inputBytes = System.Text.Encoding.UTF8.GetBytes(aSb.ToString());

            byte[] hash = md5.ComputeHash(inputBytes);

            for (int i = 0; i < hash.Length; i++)
            {
                sb.Append(hash[i].ToString("X2").ToUpper());
            }
            sb.Append(Environment.NewLine);
            sb.Append(aSb);

            return sb;
        }

        public void getPathExcel()
        {
            pathExcel = "NOK";
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Excel Files |*.xls";
            openFileDialog1.FilterIndex = 1;

            if ((this.getDirEXCEL() == "")||(this.getDirEXCEL() == "NOK"))
            {
                openFileDialog1.InitialDirectory = "C:\\";
            }
            else
            {
                openFileDialog1.InitialDirectory = this.getDirEXCEL();
            }

            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {

                pathExcel = openFileDialog1.FileName;
                lastDirExcel = Path.GetDirectoryName(pathExcel);
                this.setDirEXCEL(lastDirExcel);
                saveXml();


            }
            else
            {

                pathExcel = "NOK";
            }

        }

        public void getPathPDF()
        {
            pdfTemplate = "NOK";
            // Create an instance of the open file dialog box.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "PDF Files |*.pdf";
            openFileDialog1.FilterIndex = 1;

            if((this.getDirPDF()=="")||(this.getDirPDF() == "NOK")){
                openFileDialog1.InitialDirectory="C:\\";
            }
            else{
                openFileDialog1.InitialDirectory = this.getDirPDF();
            }
         
            openFileDialog1.Multiselect = false;

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFileDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {

                pdfTemplate = openFileDialog1.FileName;
                lastDirPDF = Path.GetDirectoryName(pdfTemplate);
                this.setDirPDF(lastDirPDF);
                saveXml();

            }
            else
            {

                pdfTemplate = "NOK";
            }

        }
        public void getPathCSV()
        {
            csvTemplate = "NOK";
            // Create an instance of the open folder dialog box.
            FolderBrowserDialog openFolderDialog1 = new FolderBrowserDialog();            
            openFolderDialog1.Description="Seleccionar carpeta para archivo .cvs";

            if ((this.getDirCSV() == "") || (this.getDirCSV() == "NOK"))
            {
                openFolderDialog1.SelectedPath = "C:\\";
            }
            else
            {
                openFolderDialog1.SelectedPath  = this.getDirCSV();
            }
            

            // Call the ShowDialog method to show the dialog box.
            DialogResult userClickedOK = openFolderDialog1.ShowDialog();

            // Process input if the user clicked OK.
            if (userClickedOK == DialogResult.OK)
            {

                csvTemplate = openFolderDialog1.SelectedPath;
                lastDirCSV = csvTemplate;
                this.setDirCSV(lastDirCSV);
                saveXml();

            }
            else
            {

                csvTemplate = "NOK";
            }

        }
        public int getMonth()
        {

            return listBox1.SelectedIndex;
        }
        public void getFields()
        {

            for (int i = 0; i < 7; i++)
            {
                Fechas[i] = 0;
            }

            //Go thru all fields in the form
            foreach (var field in form.Fields)
            {
                //Get the fields value
                string value = form.GetField(field.Key);
                int tempInt = 0;
                string[] tempFecha = new string[3];

                //#########################################
                switch (field.Key)
                {

                    case "Fecha_1":
                        if (value != "")
                        {
                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[0] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[0] = tempInt;
                        }
                        break;
                    case "Fecha_2":
                        if (value != "")
                        {
                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[1] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[1] = tempInt;
                        }
                        break;
                    case "Fecha_3":
                        if (value != "")
                        {

                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[2] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[2] = tempInt;
                        }
                        break;
                    case "Fecha_4":
                        if (value != "")
                        {

                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[3] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[3] = tempInt;
                        }
                        break;

                    case "Fecha_5":
                        if (value != "")
                        {

                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[4] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[4] = tempInt;
                        }
                        break;
                    case "Fecha_6":
                        if (value != "")
                        {

                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[5] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[5] = tempInt;
                        }
                        break;
                    case "Fecha_7":
                        if (value != "")
                        {

                            tempFecha = value.Split('/');
                            int.TryParse(tempFecha[0], out tempInt);
                            Fechas[6] = tempInt;
                            //##################################
                            int.TryParse(tempFecha[1], out tempInt);
                            Mes[6] = tempInt;
                        }
                        break;
                    //-------------------------------------------
                    case "N1":
                        hNormal[0] = value;
                        break;
                    case "N2":
                        hNormal[1] = value;
                        break;
                    case "N3":
                        hNormal[2] = value;
                        break;

                    case "N4":
                        hNormal[3] = value;
                        break;

                    case "N5":
                        hNormal[4] = value;
                        break;

                    case "N6":
                        hNormal[5] = value;
                        break;

                    case "N7":
                        hNormal[6] = value;
                        break;
                    //-------------------------------------------
                    case "E1":
                        hExtra[0] = value;
                        break;
                    case "E2":
                        hExtra[1] = value;
                        break;
                    case "E3":
                        hExtra[2] = value;
                        break;

                    case "E4":
                        hExtra[3] = value;
                        break;

                    case "E5":
                        hExtra[4] = value;
                        break;

                    case "E6":
                        hExtra[5] = value;
                        break;

                    case "E7":
                        hExtra[6] = value;
                        break;
                    //-------------------------------------------
                    case "F1":
                        hFestiva[0] = value;
                        break;
                    case "F2":
                        hFestiva[1] = value;
                        break;
                    case "F3":
                        hFestiva[2] = value;
                        break;

                    case "F4":
                        hFestiva[3] = value;
                        break;

                    case "F5":
                        hFestiva[4] = value;
                        break;

                    case "F6":
                        hFestiva[5] = value;
                        break;

                    case "F7":
                        hFestiva[6] = value;
                        break;
                    //-------------------------------------------
                    case "Z1":
                        hZombie[0] = value;
                        break;
                    case "Z2":
                        hZombie[1] = value;
                        break;
                    case "Z3":
                        hZombie[2] = value;
                        break;

                    case "Z4":
                        hZombie[3] = value;
                        break;

                    case "Z5":
                        hZombie[4] = value;
                        break;

                    case "Z6":
                        hZombie[5] = value;
                        break;

                    case "Z7":
                        hZombie[6] = value;
                        break;
                    //-------------------------------------------
                    case "V1":
                        hViaje[0] = value;
                        break;
                    case "V2":
                        hViaje[1] = value;
                        break;
                    case "V3":
                        hViaje[2] = value;
                        break;

                    case "V4":
                        hViaje[3] = value;
                        break;

                    case "V5":
                        hViaje[4] = value;
                        break;

                    case "V6":
                        hViaje[5] = value;
                        break;

                    case "V7":
                        hViaje[6] = value;
                        break;
                    //-------------------------------------------
                    case "Total_1":
                        totales[0] = value;
                        break;
                    case "Total_2":
                        totales[1] = value;
                        break;
                    case "Total_3":
                        totales[2] = value;
                        break;

                    case "Total_4":
                        totales[3] = value;
                        break;

                    case "Total_5":
                        totales[4] = value;
                        break;

                    case "Total_6":
                        totales[5] = value;
                        break;

                    case "Total_7":
                        totales[6] = value;
                        break;
                    //-------------------------------------------
                    case "Dia1":
                        Dia[0] = value;
                        break;
                    case "Dia2":
                        Dia[1] = value;
                        break;
                    case "Dia3":
                        Dia[2] = value;
                        break;

                    case "Dia4":
                        Dia[3] = value;
                        break;

                    case "Dia5":
                        Dia[4] = value;
                        break;

                    case "Dia6":
                        Dia[5] = value;
                        break;

                    case "Dia7":
                        Dia[6] = value;
                        break;
                    case "Proyecto":
                        Proyecto = value;
                        break;
                    case "Cliente":
                        Cliente = value;
                        break;
                    case "Semana Año":
                        Semana = value;
                        break;
                    default:
                        break;


                }
                //#########################################                
            }

        }
        public void getFields4CSV()
        {
            string[] temp;
            byte[] tempBytes;
            TipoIntervencion = "";

            for (int i = 0; i < 7; i++)
            {
                FechaCompleta[i] = "";
            }
            //Go thru all fields in the form
            foreach (var field in form.Fields)
            {
                //Get the fields value
                string value = form.GetField(field.Key);

                //#########################################
                switch (field.Key)
                {
                    //------Datos parte-------------------------------------                    
                    case "Proyecto":
                        Proyecto = value;
                        break;
                    case "Técnicoa":
                        Tecnico = value;
                        break;
                    case "Semana Año":
                        SemanaCSV = value;
                        SemanaCSV = SemanaCSV.Trim();
                        temp = SemanaCSV.Split('/');
                        SemanaCSV = temp[0].Trim();
                        break;                    
                    case "MotivoIntervencion": 
                        tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(value);
                        value = System.Text.Encoding.UTF8.GetString(tempBytes);
                        value =value.TrimEnd('\r', '\n');
                        MotivoIntervencion = Regex.Replace(value, @"\r\n?|\n", "\\n");
                        break;
                    case "AccionesRealizadas":
                        tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(value);
                        value = System.Text.Encoding.UTF8.GetString(tempBytes);
                        value = value.TrimEnd('\r', '\n');
                        AccionesRealizadas = Regex.Replace(value, @"\r\n?|\n", "\\n");
                        break;
                    case "Observaciones":
                        tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(value);
                        value = System.Text.Encoding.UTF8.GetString(tempBytes);
                        value = value.TrimEnd('\r', '\n');
                        Observaciones = Regex.Replace(value, @"\r\n?|\n", "\\n");
                        break;
                    case "Grupo1":
                        if (value != "")
                        {
                            if (value == "Si")
                            {
                                Estado = "Finalizado";
                            }
                            if (value == "en_curso")
                            {
                                Estado = "En curso";
                            }
                            if (value == "en_seguimiento")
                            {
                                Estado = "En seguimiento";
                            }
                            if (value == "No")
                            {
                                Estado = "No";
                            }
                        }
                        break;
                    case "Programacion":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Programacion,";
                            }
                        }
                        break;
                    case "Instalación":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Instalacion,";
                            }

                        }
                        break;
                    case "Sat":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "SAT,";
                            }

                        }
                        break;
                    case "Seguimiento":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Seguimiento,";
                            }

                        }
                        break;
                    case "Revisión":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Revision,";
                            }

                        }
                        break;
                    case "Formación":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Formacion,";
                            }

                        }
                        break;
                    case "Consultoria":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Consultoria,";
                            }

                        }
                        break;
                    case "Taller":
                        if (value != "")
                        {
                            if (value == "Sí")
                            {
                                TipoIntervencion = TipoIntervencion + "Taller,";
                            }

                        }
                        break;                    
                    //------------------------------
                    case "Fecha_1":
                        if (value != "")
                        {
                            FechaCompleta[0] = value;

                        }
                        break;
                    case "Fecha_2":
                        if (value != "")
                        {
                            FechaCompleta[1] = value;

                        }
                        break;
                    case "Fecha_3":
                        if (value != "")
                        {
                            FechaCompleta[2] = value;

                        }
                        break;
                    case "Fecha_4":
                        if (value != "")
                        {
                            FechaCompleta[3] = value;

                        }
                        break;

                    case "Fecha_5":
                        if (value != "")
                        {
                            FechaCompleta[4] = value;

                        }
                        break;
                    case "Fecha_6":
                        if (value != "")
                        {
                            FechaCompleta[5] = value;

                        }
                        break;
                    case "Fecha_7":
                        if (value != "")
                        {
                            FechaCompleta[6] = value;

                        }
                        break;                  
                    //----Tabla resumen horas--------------------
                    case "N1":
                        hNormal[0] = value;
                        break;
                    case "N2":
                        hNormal[1] = value;
                        break;
                    case "N3":
                        hNormal[2] = value;
                        break;

                    case "N4":
                        hNormal[3] = value;
                        break;

                    case "N5":
                        hNormal[4] = value;
                        break;

                    case "N6":
                        hNormal[5] = value;
                        break;

                    case "N7":
                        hNormal[6] = value;
                        break;
                    //-------------------------------------------
                    case "E1":
                        hExtra[0] = value;
                        break;
                    case "E2":
                        hExtra[1] = value;
                        break;
                    case "E3":
                        hExtra[2] = value;
                        break;

                    case "E4":
                        hExtra[3] = value;
                        break;

                    case "E5":
                        hExtra[4] = value;
                        break;

                    case "E6":
                        hExtra[5] = value;
                        break;

                    case "E7":
                        hExtra[6] = value;
                        break;
                    //-------------------------------------------
                    case "F1":
                        hFestiva[0] = value;
                        break;
                    case "F2":
                        hFestiva[1] = value;
                        break;
                    case "F3":
                        hFestiva[2] = value;
                        break;

                    case "F4":
                        hFestiva[3] = value;
                        break;

                    case "F5":
                        hFestiva[4] = value;
                        break;

                    case "F6":
                        hFestiva[5] = value;
                        break;

                    case "F7":
                        hFestiva[6] = value;
                        break;
                    //-------------------------------------------
                    case "Z1":
                        hZombie[0] = value;
                        break;
                    case "Z2":
                        hZombie[1] = value;
                        break;
                    case "Z3":
                        hZombie[2] = value;
                        break;

                    case "Z4":
                        hZombie[3] = value;
                        break;

                    case "Z5":
                        hZombie[4] = value;
                        break;

                    case "Z6":
                        hZombie[5] = value;
                        break;

                    case "Z7":
                        hZombie[6] = value;
                        break;
                    //-------------------------------------------
                    case "V1":
                        hViaje[0] = value;
                        break;
                    case "V2":
                        hViaje[1] = value;
                        break;
                    case "V3":
                        hViaje[2] = value;
                        break;

                    case "V4":
                        hViaje[3] = value;
                        break;

                    case "V5":
                        hViaje[4] = value;
                        break;

                    case "V6":
                        hViaje[5] = value;
                        break;

                    case "V7":
                        hViaje[6] = value;
                        break;
                    //-------------------------------------------
                    case "Total_1":
                        totales[0] = value;
                        break;
                    case "Total_2":
                        totales[1] = value;
                        break;
                    case "Total_3":
                        totales[2] = value;
                        break;

                    case "Total_4":
                        totales[3] = value;
                        break;

                    case "Total_5":
                        totales[4] = value;
                        break;

                    case "Total_6":
                        totales[5] = value;
                        break;

                    case "Total_7":
                        totales[6] = value;
                        break;
                    //-----TRAMAS HORARIAS--------------------------------------
                    //-----HORA SALIDA 1--------------------------------------
                    case "HS1":
                        hsp[0] = value;
                        break;
                    case "HS2":
                        hsp[1] = value;
                        break;
                    case "HS3":
                        hsp[2] = value;
                        break;
                    case "HS4":
                        hsp[3] = value;
                        break;
                    case "HS5":
                        hsp[4] = value;
                        break;
                    case "HS6":
                        hsp[5] = value;
                        break;
                    case "HS7":
                        hsp[6] = value;
                        break;
                    //-----HORA LLEGADA 1--------------------------------------
                    case "HL1":
                        hlp[0] = value;
                        break;
                    case "HL2":
                        hlp[1] = value;
                        break;
                    case "HL3":
                        hlp[2] = value;
                        break;
                    case "HL4":
                        hlp[3] = value;
                        break;
                    case "HL5":
                        hlp[4] = value;
                        break;
                    case "HL6":
                        hlp[5] = value;
                        break;
                    case "HL7":
                        hlp[6] = value;
                        break;
                    //-------HORA INICIO 1------------------------------------
                    case "HI1":
                        hip[0] = value;
                        break;
                    case "HI2":
                        hip[1] = value;
                        break;
                    case "HI3":
                        hip[2] = value;
                        break;
                    case "HI4":
                        hip[3] = value;
                        break;
                    case "HI5":
                        hip[4] = value;
                        break;
                    case "HI6":
                        hip[5] = value;
                        break;
                    case "HI7":
                        hip[6] = value;
                        break;
                    //-------HORA FIN 1------------------------------------
                    case "HF1":
                        hfp[0] = value;
                        break;
                    case "HF2":
                        hfp[1] = value;
                        break;
                    case "HF3":
                        hfp[2] = value;
                        break;
                    case "HF4":
                        hfp[3] = value;
                        break;
                    case "HF5":
                        hfp[4] = value;
                        break;
                    case "HF6":
                        hfp[5] = value;
                        break;
                    case "HF7":
                        hfp[6] = value;
                        break;
                    //-------HORA INICIO 2------------------------------------
                    case "HI11":
                        his[0] = value;
                        break;
                    case "HI12":
                        his[1] = value;
                        break;
                    case "HI13":
                        his[2] = value;
                        break;
                    case "HI14":
                        his[3] = value;
                        break;
                    case "HI15":
                        his[4] = value;
                        break;
                    case "HI16":
                        his[5] = value;
                        break;
                    case "HI17":
                        his[6] = value;
                        break;
                    //-------HORA FIN 2------------------------------------
                    case "HF11":
                        hfs[0] = value;
                        break;
                    case "HF12":
                        hfs[1] = value;
                        break;
                    case "HF13":
                        hfs[2] = value;
                        break;
                    case "HF14":
                        hfs[3] = value;
                        break;
                    case "HF15":
                        hfs[4] = value;
                        break;
                    case "HF16":
                        hfs[5] = value;
                        break;
                    case "HF17":
                        hfs[6] = value;
                        break;
                    //------HORA SALIDA 2-------------------------------------
                    case "HS11":
                        hss[0] = value;
                        break;
                    case "HS12":
                        hss[1] = value;
                        break;
                    case "HS13":
                        hss[2] = value;
                        break;
                    case "HS14":
                        hss[3] = value;
                        break;
                    case "HS15":
                        hss[4] = value;
                        break;
                    case "HS16":
                        hss[5] = value;
                        break;
                    case "HS17":
                        hss[6] = value;
                        break;
                    //------HORA LLEGADA 2-------------------------------------
                    case "HL11":
                        hls[0] = value;
                        break;
                    case "HL12":
                        hls[1] = value;
                        break;
                    case "HL13":
                        hls[2] = value;
                        break;
                    case "HL14":
                        hls[3] = value;
                        break;
                    case "HL15":
                        hls[4] = value;
                        break;
                    case "HL16":
                        hls[5] = value;
                        break;
                    case "HL17":
                        hls[6] = value;
                        break;          
                    default:
                        break;

                }
                //#########################################
            }

        }
        private void sendAttachment()
        {
            CreateMailItem();
        }
        private void CreateMailItem()
        {
                string smtpAddress = "armrobotics.com";
                try
                {
                    Outlook.Application outlookApp = new Outlook.Application();
                    Outlook.Accounts accounts = outlookApp.Session.Accounts;
                    Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
                    
                    
                    foreach (Outlook.Account account in accounts)
                    {
                        // When the e-mail address matches, send the mail.                         
                        if (account.SmtpAddress.Contains(smtpAddress))
                        {
                            Outlook.Recipient recipient = outlookApp.Session.CreateRecipient(account.SmtpAddress);
                            mailItem.Sender = recipient.AddressEntry;
                            mailItem.SendUsingAccount = account;
                    }
                    }
                    mailItem.Subject = "Envío CSV " + nameFileCSV;
                    mailItem.To = getEmail();
                    if (DateTime.Now.ToString("tt", CultureInfo.InvariantCulture) == "AM")
                    {
                        mailItem.Body = "Buenos días\n\nAdjunto archivos\n\n CSV: " + nameFileCSV + "\n\n PDF vinculado: " + Path.GetFileName(pdfTemplate) + "\n\nUn saludo";
                    }
                    else
                    {
                        mailItem.Body = "Buenos tardes\n\nAdjunto archivos\n\n CSV: " + nameFileCSV + "\n\n PDF vinculado: " + Path.GetFileName(pdfTemplate) + "\n\nUn saludo";
                }                    
                    mailItem.Importance = Outlook.OlImportance.olImportanceNormal;
                    mailItem.Display(false);
                    if (filePathCSV.Length > 0)
                    {
                        mailItem.Attachments.Add(
                            filePathCSV,
                            Outlook.OlAttachmentType.olByValue,
                            1,
                            nameFileCSV);
                        mailItem.Attachments.Add(
                            pdfTemplate,
                            Outlook.OlAttachmentType.olByValue,
                            1,
                            Path.GetFileName(pdfTemplate));

                }
                    else
                    {
                        MessageBox.Show("Error: No se adjunta archivo.\n CSV con longitud 0\n");

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: ocurrio un error durante el envio del archivo\n\n" + ex);
                }           

        }

        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                {
                    using (var stream = client.OpenRead("http://www.google.com"))
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }
        public void getFormPdf()
        {
            //========================================            
            pdfReader = new PdfReader(pdfTemplate);
            //Get the Form
            form = pdfReader.AcroFields;

        }
        public void getExcelCell()
        {

            currentSheet = "Parte de horas";
            xls = new _Excel.Application();
            usp = xls.DecimalSeparator;

            wkbk = xls.Workbooks.Open(pathExcel,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing);

            sheet = wkbk.Worksheets;
            excelWorksheet = (_Excel.Worksheet)sheet.get_Item(currentSheet);


        }
        //Save xml file
        public static void saveXml()
        {
            doc1.Save(pathXml);
        }
        
        //Setters
        public static void setCodOper(String aSTR)
        {

            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/CODOPER");
            node.InnerText = aSTR;

        }
        public static void setEmail(String aSTR)
        {

            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/EMAIL");
            node.InnerText = aSTR;

        }
        private void setDirPDF(String aSTR)
        {

            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIRPDF");
            node.InnerText = aSTR;

        }
        private void setDirEXCEL(String aSTR)
        {

            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIREXCEL");
            node.InnerText = aSTR;

        }
        private void setDirCSV(String aSTR)
        {

            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIRCSV");
            node.InnerText = aSTR;

        }
        //Getters
        public static String getCodOper()
        {
            string attr = null;
            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/CODOPER");
            attr = node.InnerText;

            return attr;
        }
        public static String getEmail()
        {
            string attr = null;
            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/EMAIL");
            attr = node.InnerText;

            return attr;
        }
        private String getDirPDF()
        {
            string attr = null;
            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIRPDF");
            attr = node.InnerText;

            return attr;
        }
        private String getDirEXCEL()
        {
            string attr = null;
            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIREXCEL");
            attr = node.InnerText;

            return attr;
        }
        private String getDirCSV()
        {
            string attr = null;
            XmlNode node = doc1.DocumentElement.SelectSingleNode("/USER/DIRCSV");
            attr = node.InnerText;

            return attr;
        }
        public void writeCell(String acol, String arow, String value)
        {

            excelCell = (_Excel.Range)excelWorksheet.get_Range(acol, arow);
            excelCell.Value2 = value;

        }
        public string readCell(String acol, String arow)
        {
            string value;
            excelCell = (_Excel.Range)excelWorksheet.get_Range(acol, arow);
            value = excelCell.Text;           

            return value;

        }
        public void prepareData()
        {          
            string[] temp2;

            for (int i = 0; i < 7; i++)
            {
                if ((Fechas[i] != 0) && (Mes[i] == month))
                {
                    //Semana
                    Semana = Semana.Trim();
                    temp2 = Semana.Split('/');
                    Semanas[i] = temp2[0];
                    //Cliente                    
                    Clientes[i] = Cliente;
                    //Proyecto
                    Proyectos[i] = Proyecto;
                    //Horas normales                  
                    double.TryParse(hNormal[i], System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out dHNormal[i]);
                    //Horas zombie
                    double.TryParse(hZombie[i], System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out dHZombie[i]);
                    //Horas extra
                    double.TryParse(hExtra[i], System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out dHExtra[i]);
                    //Horas festivas
                    double.TryParse(hFestiva[i], System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out dHFestiva[i]);
                    //Horas viajes
                    double.TryParse(hViaje[i], System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out dHViaje[i]);
                    //========================================================================================
                    dHSumas[i] = dHNormal[i] + dHZombie[i] + dHExtra[i] + dHFestiva[i] + dHViaje[i];
                    //======================================================================================== 
                }
            }

        }
        public void setCells()
        {

            string temp = null;
            double temp1 = 0.0;


            for (int i = 0; i < 7; i++)
            {

                temp1 = 0.0;
                if ((Fechas[i] != 0) && (Mes[i] == month))
                {
                    //Semana
                    temp = "C" + (Fechas[i] + 8).ToString();
                    this.writeCell(temp, temp, Semanas[i]);

                    //Cliente                    
                    temp = "E" + (Fechas[i] + 8).ToString();
                    this.writeCell(temp, temp, Clientes[i]);

                    //Proyecto
                    temp = "D" + (Fechas[i] + 8).ToString();
                    this.writeCell(temp, temp, Proyectos[i]);


                    //========================================================================================
                    switch (CheckDia(Dia[i]))
                    {
                        case 1:
                            if ((dHZombie[i] <= 0) && (dHFestiva[i] <= 0))
                            {
                                if (dHSumas[i] < Dia_Normal)
                                {
                                    temp1 = (Dia_Normal - dHSumas[i]);
                                    subSetCells(Fechas[i], dHNormal[i], 0, 0, 0, temp1, 0);
                                }
                                else
                                {
                                    subSetCells(Fechas[i], Dia_Normal, 0, dHExtra[i], 0, 0, dHViaje[i]);

                                }
                            }
                            else if (dHFestiva[i] > 0)
                            {
                                subSetCells(Fechas[i], 0, 0, 0, dHFestiva[i], 0, 0);
                            }
                            else if (dHZombie[i] > 0)
                            {
                                if (dHSumas[i] < Horas_Zombie)
                                {
                                    temp1 = (Horas_Zombie - dHSumas[i]);
                                    subSetCells(Fechas[i], 0, dHZombie[i], 0, 0, temp1, 0);
                                }
                                else
                                {
                                    subSetCells(Fechas[i], 0, Horas_Zombie, dHExtra[i], 0, 0, dHViaje[i]);

                                }
                            }

                            break;
                        case 2:
                            if ((dHZombie[i] <= 0) && (dHFestiva[i] <= 0))
                            {
                                if (dHSumas[i] < 5.0)
                                {
                                    temp1 = (5.0 - dHSumas[i]);
                                    subSetCells(Fechas[i], dHNormal[i], 0, 0, 0, temp1, 0);
                                }
                                else
                                {
                                    subSetCells(Fechas[i], 5, 0, dHExtra[i], 0, 0, dHViaje[i]);

                                }
                            }
                            else if (dHFestiva[i] > 0)
                            {
                                subSetCells(Fechas[i], 0, 0, 0, dHFestiva[i], 0, 0);
                            }
                            else if (dHZombie[i] > 0)
                            {
                                subSetCells(Fechas[i], 0, 0, 0, dHFestiva[i], 0, 0);
                            }
                            break;
                        case 3:
                            subSetCells(Fechas[i], 0, 0, 0, dHFestiva[i], 0, 0);
                            break;
                        default:
                            break;
                    }
                    //========================================================================================

                }
            }
        }
        public void subSetCells(int aFecha, double aHN, double aHZ, double aHE, double aHF, double aHVN, double aHVP)
        {
            String temp;

            //Horas normales
            temp = "G" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHN == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHN.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }
                
            }
            else
            {
                if (aHN == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHN.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }
                
            }
            //Horas zombie
            temp = "I" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHZ == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHZ.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }
                
            }
            else
            {
                if (aHZ == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHZ.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }
                
            }
            //Horas extra
            temp = "J" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHE == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            else
            {
                if (aHE == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHE.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            //Horas festivas
            temp = "K" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHF == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHF.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            else
            {
                if (aHF == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHF.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            //Horas viajes negativas
            temp = "L" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHVN == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHVN.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            else
            {
                if (aHVN == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHVN.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            //Horas viajes positivas
            temp = "M" + (aFecha + 8).ToString();
            if (usp == ",")
            {
                if (aHVP == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHVP.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }
            else
            {
                if (aHVP == 0)
                {
                    this.writeCell(temp, temp, "");
                }
                else
                {
                    this.writeCell(temp, temp, (aHVP.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)));
                }

            }

        }  
        private int CheckDia(String aDia)
        {
            int result = 0;

            if ((aDia == "L") || (aDia == "M") || (aDia == "X") || (aDia == "J"))
            {
                result = 1;
            }

            if (aDia == "V")
            {
                result = 2;
            }
            if ((aDia == "S") || (aDia == "D"))
            {
                result = 3;
            }

            return result;
        }        
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Se va a proceder a export datos del pdf a excel\n ¿Quiere continuar?", "Atencion",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                if ((pathExcel != "NOK") && (pdfTemplate != "NOK") && (textBox1.Text != "") && (textBox2.Text != "")){

                    if (!IsOpened(pathExcel))
                    {
                        this.export();
                    }
                    else
                    {
                        MessageBox.Show(" Error: El archivo excel de destino se encuentra abierto.\n Cierre el archivo antes de proceder! ");
               
                    }

                }
                else
                {
                    MessageBox.Show(" Error: Ruta de archivos incorrecta.\n Utilice los botones de seleccion de ruta de archivos:\n * Origen PDF\n * Destino Excel ");
                }
                
                 
                
                
            }
            else if (result == DialogResult.No)
            {
                //code for No
            } 
            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Se va a proceder a export datos del pdf a archivo CSV\n ¿Quiere continuar?", "Atencion",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                if ((pdfTemplate != "NOK") && (textBox1.Text != ""))
                {
                    if ((getCodOper() != "") && (getEmail() != ""))
                    {
                        this.getPathCSV();
                        this.exportCSV();
                    }
                    else
                    {
                        MessageBox.Show(" Error: Introduzca su codigo de trabajador y direccion email de envio ");
                    }
                    
                }
                else
                {

                    MessageBox.Show(" Error: Ruta de archivos incorrecta.\n Utilice los botones de seleccion de ruta de archivos:\n * Origen PDF\n  ");

                }
            }
            else if (result == DialogResult.No)
            {
                //code for No
            } 
        }

        

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

        }

        private void configuracionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 frm = new GUI_PDF2EXCEL.Form2();
            frm.ShowDialog();
        }
    }
}
    
    
       

