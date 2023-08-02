using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;
using FirebirdSql.Data.FirebirdClient;
using OfficeOpenXml;
using SpreadsheetLight;
using SpreadsheetLight.Charts;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Xml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Admin_Report
{
    public partial class Form1 : Form
    {
        List<(string, decimal)> datosFilas = new List<(string, decimal)>();
        List<(string, decimal)> datosFilasS = new List<(string, decimal)>();
        List<(string, decimal)> datosFacturistas = new List<(string, decimal)>();
        List<int> id = new List<int> { };
        List<string> id_new = new List<string> { };
        List<string> folio = new List<string> { };
        List<string> folio_fact_new = new List<string> { };
        List<string> folio_fact_new2 = new List<string> { };
        List<string> surtidor = new List<string> { };
        List<string> horaInicio = new List<string> { };
        List<decimal> importe = new List<decimal> { };
        List<string> horaFin = new List<string> { };
        List<string> duracion = new List<string> { };
        List<int> renglones = new List<int> { };
        List<string> fecha = new List<string> { };
        List<decimal> importenew = new List<decimal> { };
        List<string> folionew = new List<string> { };
        List<int> renglonesnew = new List<int> { };
        List<int> renglonesnew2 = new List<int> { };
        List<int> Renglones_past = new List<int> { };
        List<int> Pedidos_past = new List<int> { };
        List<double> Minutos_past = new List<double> { };
        List<decimal> Importe_past = new List<decimal> { };
        List<string> Surtidor_list = new List<string> { };
        string Buscar_Folio;
        int Num_renglones;
        string filePath;
        string oficial_id;
        Decimal Bono, DineroGenerado;
        SLStyle style = new SLStyle();
        SLStyle style2 = new SLStyle();
        SLStyle style3 = new SLStyle();
        SLStyle style4 = new SLStyle();
        SLStyle style5 = new SLStyle();
        SLStyle style6 = new SLStyle();
        public Form1()
        {
            InitializeComponent();
            style.Fill.SetPattern(PatternValues.LightGray, System.Drawing.Color.LightBlue, System.Drawing.Color.Empty);
            style2.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightSteelBlue, System.Drawing.Color.Empty);
            style3.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.Beige, System.Drawing.Color.Empty);
            style4.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.BurlyWood, System.Drawing.Color.Empty);
            style5.Fill.SetPattern(PatternValues.Solid, System.Drawing.Color.LightYellow, System.Drawing.Color.Empty);
            style6.Fill.SetPattern(PatternValues.LightGray, System.Drawing.Color.LightBlue, System.Drawing.Color.Empty);
            style6.FormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
            style3.FormatCode = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
            style.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style.Border.TopBorder.Color = System.Drawing.Color.Black;
            style2.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style2.Border.TopBorder.Color = System.Drawing.Color.Black;
            style3.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style3.Border.TopBorder.Color = System.Drawing.Color.Black;
            style4.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style4.Border.TopBorder.Color = System.Drawing.Color.Black;
            style5.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style5.Border.TopBorder.Color = System.Drawing.Color.Black;
            style6.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            style6.Border.TopBorder.Color = System.Drawing.Color.Black;
            style.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style.Border.RightBorder.Color = System.Drawing.Color.Black;
            style2.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style2.Border.RightBorder.Color = System.Drawing.Color.Black;
            style3.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style3.Border.RightBorder.Color = System.Drawing.Color.Black;
            style4.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style4.Border.RightBorder.Color = System.Drawing.Color.Black;
            style5.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style5.Border.RightBorder.Color = System.Drawing.Color.Black;
            style6.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            style6.Border.RightBorder.Color = System.Drawing.Color.Black;
            style.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style2.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style2.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style3.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style3.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style4.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style4.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style5.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style5.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style6.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            style6.Border.LeftBorder.Color = System.Drawing.Color.Black;
            style.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style.Border.BottomBorder.Color = System.Drawing.Color.Black;
            style2.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style2.Border.BottomBorder.Color = System.Drawing.Color.Black;
            style3.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style3.Border.BottomBorder.Color = System.Drawing.Color.Black;
            style4.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style4.Border.BottomBorder.Color = System.Drawing.Color.Black;
            style5.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style5.Border.BottomBorder.Color = System.Drawing.Color.Black;
            style6.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;
            style6.Border.BottomBorder.Color = System.Drawing.Color.Black;
        }
        public void EstadisticasFinales()   
        {
            SLDocument sl = new SLDocument(filePath);
            // Agregar una nueva hoja
            sl.AddWorksheet("Estadistica Final");
            sl.SelectWorksheet("Estadistica Final");
            int[] columnas = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 };
            foreach (int columna in columnas)
            {
                sl.SetColumnWidth(columna, 30);

            }
            sl.SetCellValue("A1", "SURTIDOR");
            sl.SetCellValue("B1", "PEDIDOS SOLICITADOS");
            sl.SetCellValue("C1", "IMPORTE SOLICITADO");
            sl.SetCellValue("D1", "RENGLONES SOLICITADOS");
            sl.SetCellValue("E1", "MINUTOS EN SURTIR");
            sl.SetCellValue("F1", "RENGLONES SURTIDOS");
            sl.SetCellValue("G1", "IMPORTE SURTIDO");
            sl.SetCellValue("H1", "IMPORTE TOTAL SOLICITADO");
            sl.SetCellValue("I1", "IMPORTE TOTAL SURTIDO");
            sl.SetCellValue("J1", "RENGLONES TOTALES SOLICITADOS");
            sl.SetCellValue("K1", "RENGLONES TOTALES SURTIDOS");
            sl.SetCellValue("L1", "PORCENTAJE DE RENGLONES SURTIDOS");
            sl.SetCellValue("M1", "PORCENTAJE DE IMPORTE SURTIDO");
            sl.SetCellValue("N1", "PORCENTAJE IMPORTE / RENGLONES");
            for (int i = 2; i < Surtidor_list.Count + 2; i++)
            {
                sl.SetCellValue("A" + i, Surtidor_list[i - 2]);
                sl.SetCellValue("B" + i, Pedidos_past[i - 2]);
                sl.SetCellValue("C" + i, Importe_past[i - 2]);
                sl.SetCellValue("D" + i, Renglones_past[i - 2]);
                sl.SetCellValue("E" + i, Minutos_past[i - 2]);

            }
            int k = 2;
            for (int i = 0; i < folio.Count; i++)
            {
                while (sl.HasCellValue("A" + k))
                {
                    if (sl.GetCellValueAsString(k, 1) == surtidor[i])
                    {
                        decimal imp = sl.GetCellValueAsDecimal("G" + k);
                        int renglon_soli = sl.GetCellValueAsInt32("F" + k);
                        sl.SetCellValue("G" + k, importenew[i] + imp);
                        sl.SetCellValue("F" + k, renglonesnew[i] + renglon_soli);
                        k = 2;
                        break;
                    }
                    k++;
                }
            }
            while (sl.HasCellValue("A" + k))
            {
                decimal imp_soli = sl.GetCellValueAsDecimal("C" + k);
                decimal impo_soli_1 = sl.GetCellValueAsDecimal("H2");
                decimal imp_sur = sl.GetCellValueAsDecimal("G" + k);
                decimal impo_sur_1 = sl.GetCellValueAsDecimal("I2");
                int ren_soli = sl.GetCellValueAsInt32("D" + k);
                int ren_soli_1 = sl.GetCellValueAsInt32("J2");
                decimal ren_sur = sl.GetCellValueAsDecimal("F" + k);
                decimal ren_sur_1 = sl.GetCellValueAsDecimal("K2");
                sl.SetCellValue("H2", imp_soli + impo_soli_1);
                sl.SetCellValue("I2", imp_sur + impo_sur_1);
                sl.SetCellValue("J2", ren_soli + ren_soli_1);
                sl.SetCellValue("K2", ren_sur + ren_sur_1);
                k++;
            }
            k = 2;
            while (sl.HasCellValue("A" + k))
            {
                decimal ren_sur = sl.GetCellValueAsDecimal("F" + k);
                decimal ren_sur_1 = sl.GetCellValueAsDecimal("K2");
                sl.SetCellValue("L" + k, 100 * (ren_sur / ren_sur_1));
                decimal imp_sur = sl.GetCellValueAsDecimal("G" + k);
                decimal imp_sur_1 = sl.GetCellValueAsDecimal("I2");
                sl.SetCellValue("M" + k, 100 * (imp_sur / imp_sur_1));
                sl.SetCellValue("N" + k, (sl.GetCellValueAsDecimal("M" + k) + sl.GetCellValueAsDecimal("L" + k)) / 2);
                k++;
            }
            SLChart chart;
            double fChartHeight = 40.0;
            double fChartWidth = 20;
            // Crear un gráfico de barras en la celda C1
            int filas = sl.GetWorksheetStatistics().NumberOfRows;
            chart = sl.CreateChart("A1", "G" + filas);
            chart.SetChartType(SLColumnChartType.ClusteredColumn);
            chart.SetChartStyle(SLChartStyle.Style14);
            sl.AddWorksheet("Graficas");
            sl.SelectWorksheet("Graficas");
            chart.SetChartPosition(0, 0, 0 + fChartHeight, 0 + fChartWidth);
            sl.InsertChart(chart);
            // Guardar el archivo con la nueva hoja agregada
            sl.SaveAs(filePath);
            //FINAL REPORT
            using (SLDocument nDoc = new SLDocument(filePath))
            {

                nDoc.SelectWorksheet("Estadistica Final");
                int numFilas = nDoc.GetWorksheetStatistics().NumberOfRows;
                // Definimos el rango de la columna A que queremos ordenar (ejemplo: de la fila 2 a la 11)
                int filaInicio = 2;
                int filaFin = numFilas;

                // Creamos una lista para almacenar los datos de las filas y la columna A

                // Leer los datos de las filas y la columna A
                for (int fila = filaInicio; fila <= filaFin; fila++)
                {
                    string valorColumnaA = nDoc.GetCellValueAsString(fila, 1);
                    decimal valorColumnaN = nDoc.GetCellValueAsDecimal(fila, 14);
      
                    string subporcentaje = valorColumnaN.ToString("0.00");
                    datosFilasS.Add((valorColumnaA,decimal.Parse(subporcentaje)));
                }
                datosFilasS.Sort();

                // Guardar los cambios en el archivo ordenado
                nDoc.SaveAs(filePath);
            }
            SLDocument doc = new(filePath);
            doc.AddWorksheet("Final Report");
            doc.SelectWorksheet("Final Report");
            int[] columnas2 = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };
            foreach (int columna in columnas)
            {
                doc.SetColumnWidth(columna, 30);
                if (columna == 1)
                    doc.SetColumnWidth(columna, 50);
            }
            int p = 4;
            doc.SetCellValue("A3", "EMPLEADO");
            doc.SetCellValue("C3", "PORCENTAJE DE SURTIDO");

            foreach (var tupla in datosFilasS)
            {
                doc.SetCellValue("A" + p, tupla.Item1);
                doc.SetCellValue("C" + p, tupla.Item2);

                ++p;
            }
            doc.SaveAs(filePath);
        }


        private void Btn_Surtido_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;";

            // Show the dialog and check if the user clicked the OK button
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                SLDocument sl = new(selectedFileName);
                sl.SelectWorksheet("Reporte Surtido");
                int numFilas = sl.GetWorksheetStatistics().NumberOfRows;
                for (int i = 2; i <= numFilas; ++i)
                {
                    renglonesnew.Add(0);
                    importenew.Add(0);
                    folionew.Add("0");
                    string Buscar_Folio = sl.GetCellValueAsString(i, 2);
                    if (Buscar_Folio[1] == 'O' || Buscar_Folio[1] == 'E')
                    {
                        int cont = 9 - Buscar_Folio.Length;
                        string prefix = Buscar_Folio.Substring(0, 2);
                        string suffix = Buscar_Folio.Substring(2);
                        string patch = "";
                        for (int j = 0; j < cont; j++)
                        {
                            patch = patch + "0";
                        }
                        Buscar_Folio = prefix + patch + suffix;
                    }
                    else if (Buscar_Folio[0] == 'P')
                    {
                        int cont = 9 - Buscar_Folio.Length;
                        string prefix = Buscar_Folio.Substring(0, 1);
                        string suffix = Buscar_Folio.Substring(1);
                        string patch = "";
                        for (int j = 0; j < cont; j++)
                        {
                            patch = patch + "0";
                        }
                        Buscar_Folio = prefix + patch + suffix;
                    }
                    id.Add(sl.GetCellValueAsInt32(i, 1));
                    folio.Add(Buscar_Folio);
                    surtidor.Add(sl.GetCellValueAsString(i, 3));
                    horaInicio.Add(sl.GetCellValueAsString(i, 4));
                    importe.Add(sl.GetCellValueAsDecimal(i, 5));
                    horaFin.Add(sl.GetCellValueAsString(i, 6));
                    duracion.Add(sl.GetCellValueAsString(i, 7));
                    renglones.Add(sl.GetCellValueAsInt32(i, 8));
                    fecha.Add(sl.GetCellValueAsString(i, 9));
                }
                sl.SelectWorksheet("Promedio");
                int numFilas2 = sl.GetWorksheetStatistics().NumberOfRows;
                for (int j = 2; j <= numFilas2; ++j)
                {
                    Renglones_past.Add(sl.GetCellValueAsInt32(j, 6));
                    Importe_past.Add(sl.GetCellValueAsDecimal(j, 3));
                    Surtidor_list.Add(sl.GetCellValueAsString(j, 1));
                    Pedidos_past.Add(sl.GetCellValueAsInt32(j, 2));
                    Minutos_past.Add(sl.GetCellValueAsDouble(j, 5));
                }
                FbConnection con = new FbConnection("User=SYSDBA;" + "Password=C0r1b423;" + "Database=D:\\Microsip datos\\PAPELERIA CORIBA CORNEJO.fdb;" + "DataSource=192.168.0.11;" + "Port=3050;" + "Dialect=3;" + "Charset=UTF8;");
                try
                {
                    con.Open();
                    string query0 = "SELECT * FROM DOCTOS_VE WHERE (FOLIO LIKE 'P%') AND TIPO_DOCTO = 'P' ORDER BY FECHA DESC";
                    FbCommand command0 = new FbCommand(query0, con);
                    FbDataReader reader0 = command0.ExecuteReader();
                    progressBar1.Minimum = 0; // Valor mínimo del ProgressBar
                    progressBar1.Maximum = folio.Count; // Valor máximo del ProgressBar
                    progressBar1.Step = 1; // Incremento del ProgressBar por cada iteración
                    Loading.Visible = true;
                    progressBar1.Visible = true;
                    Loading.Update();
                    int cont = 0;
                    while (reader0.Read())
                    {
                        string columna0 = reader0.GetString(0);
                        string columna1 = reader0.GetString(4);
                        decimal importes = reader0.GetDecimal(26);

                        if (folio.Contains(columna1))
                        {
                            int posicion = folio.FindIndex(folio => folio.Contains(columna1));
                            folionew[posicion] = columna1;
                            importenew[posicion] = importes;
                            //MessageBox.Show(columna1);
                            //MessageBox.Show(importe.ToString());
                            FbCommand command = new FbCommand("CALC_TOTALES_DOCTO_VE", con);
                            command.CommandType = CommandType.StoredProcedure;

                            // Parámetro V_DOCTO_VE_ID
                            FbParameter paramV_DOCTO_VE_ID = new FbParameter("V_DOCTO_VE_ID", FbDbType.Integer);
                            paramV_DOCTO_VE_ID.Value = columna0;
                            command.Parameters.Add(paramV_DOCTO_VE_ID);

                            // Parámetro V_SOLO_LECTURA
                            FbParameter paramV_SOLO_LECTURA = new FbParameter("V_SOLO_LECTURA", FbDbType.Char);
                            paramV_SOLO_LECTURA.Size = 1;
                            paramV_SOLO_LECTURA.Value = 'S';
                            command.Parameters.Add(paramV_SOLO_LECTURA);

                            // Parámetro de retorno NUM_RENGLONES
                            FbParameter paramNUM_RENGLONES = new FbParameter();
                            paramNUM_RENGLONES.ParameterName = "NUM_RENGLONES";
                            paramNUM_RENGLONES.Direction = ParameterDirection.Output;
                            paramNUM_RENGLONES.FbDbType = FbDbType.Integer;
                            command.Parameters.Add(paramNUM_RENGLONES);
                            command.ExecuteNonQuery();

                            // Obtener el valor de retorno NUM_RENGLONES
                            int Num_renglones = (int)paramNUM_RENGLONES.Value;
                            renglonesnew[posicion] = Num_renglones;
                            //final

                            progressBar1.Value = cont + 1;
                            progressBar1.Update();
                            //System.Threading.Thread.Sleep(100);
                            cont++;
                            if (cont == folio.Count)
                                break;
                        }


                    }
                    reader0.Close();
                    saveFileDialog1.Title = "";
                    saveFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
                    saveFileDialog1.FileName = "Nombre";

                    // Mostrar el cuadro de diálogo
                    DialogResult result = saveFileDialog1.ShowDialog();

                    // Verificar si se ha seleccionado un archivo y se ha hecho clic en el botón "Guardar"
                    if (result == DialogResult.OK)
                    {
                        filePath = saveFileDialog1.FileName;
                        SLDocument documento = new SLDocument();
                        string actual = documento.GetCurrentWorksheetName();
                        documento.RenameWorksheet(actual, "Estadisticas Surtido");
                        int[] columnas = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11 };
                        foreach (int columna in columnas)
                        {
                            documento.SetColumnWidth(columna, 20);
                            if (columna == 3)
                                documento.SetColumnWidth(columna, 30);
                        }
                        documento.SetCellValue("A1", "ID");
                        documento.SetCellValue("B1", "FOLIO");
                        documento.SetCellValue("C1", "SURTIDOR");
                        documento.SetCellValue("D1", "HORA DE INICIO");
                        documento.SetCellValue("E1", "IMPORTE INICIAL");
                        documento.SetCellValue("F1", "HORA DE FIN");
                        documento.SetCellValue("G1", "DURACION");
                        documento.SetCellValue("H1", "RENGLONES SOLICITADOS");
                        documento.SetCellValue("I1", "FECHA");
                        documento.SetCellValue("J1", "IMPORTE FINAL");
                        documento.SetCellValue("K1", "RENGLONES FINALES");
                        for (int g = 2; g < folio.Count + 2; g++)
                        {
                            documento.SetCellValue("A" + g, id[g - 2]);
                            documento.SetCellValue("B" + g, folio[g - 2]);
                            documento.SetCellValue("C" + g, surtidor[g - 2]);
                            documento.SetCellValue("D" + g, horaInicio[g - 2]);
                            documento.SetCellValue("E" + g, importe[g - 2]);
                            documento.SetCellValue("F" + g, horaFin[g - 2]);
                            documento.SetCellValue("G" + g, duracion[g - 2]);
                            documento.SetCellValue("H" + g, renglones[g - 2]);
                            documento.SetCellValue("I" + g, fecha[g - 2]);
                            documento.SetCellValue("J" + g, importenew[g - 2]);
                            documento.SetCellValue("K" + g, renglonesnew[g - 2]);
                        }

                        // Guardar el documento en la ubicación seleccionada por el usuario
                        documento.SaveAs(filePath);
                        progressBar1.Visible = false;
                        Loading.Visible = false;
                        EstadisticasFinales();
                        MessageBox.Show("Cargado con Éxito");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void Btn_Facturacion_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;";

            // Show the dialog and check if the user clicked the OK button
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;
                SLDocument sl = new(selectedFileName);
                int numFilas = (sl.GetWorksheetStatistics().NumberOfRows);
                List<string> Facturistas = new List<string>();
                //sl.SetCellValue(2, 5, $"=SUM(J8:J{numFilas})");
                sl.SetColumnWidth(16, 30);
                sl.SetColumnWidth(17, 30);
                sl.SetColumnWidth(18, 30);
                sl.SetCellValue(1, 16, "SUMA DE VENTAS");
                sl.SetCellValue(1, 17, "BONO CORRESPONDIENTE");
                sl.SetCellValue(1, 18, "TOTAL RENGLONES");
                string filePath2 = "C:\\BlackList\\blacklist.txt";
                string[] nombresBl = File.ReadAllLines(filePath2);
                int l = 0;

                for (int i = 8; i < numFilas; ++i)
                {

                    if (sl.HasCellValue(i, 7))
                    {
                        Buscar_Folio = sl.GetCellValueAsString(i, 7);
                        int conta = 9 - Buscar_Folio.Length;
                        string prefix = Buscar_Folio.Substring(0, 2);
                        string suffix = Buscar_Folio.Substring(2);
                        string patch = "";
                        for (int j = 0; j < conta; j++)
                        {
                            patch = patch + "0";
                        }
                        Buscar_Folio = prefix + patch + suffix;
                        folio_fact_new.Add(Buscar_Folio);
                        folio_fact_new2.Add("0");
                        id_new.Add("0");
                        renglonesnew2.Add(0);
                        //MessageBox.Show(folio_fact_new[l]);
                        l++;

                    }
                }


                for (int i = 8; i < numFilas; i++)
                {
                    if (!Facturistas.Contains(sl.GetCellValueAsString(i, 1)) && sl.GetCellValueAsString(i, 1) != "" && !nombresBl.Contains(sl.GetCellValueAsString(i, 1)))
                    {
                        Facturistas.Add(sl.GetCellValueAsString("A" + i));
                    }
                }
                Facturistas.Sort();
                FbConnection con = new FbConnection("User=SYSDBA;" + "Password=C0r1b423;" + "Database=D:\\Microsip datos\\PAPELERIA CORIBA CORNEJO.fdb;" + "DataSource=192.168.0.11;" + "Port=3050;" + "Dialect=3;" + "Charset=UTF8;");
                try
                {
                    progressBar3.Minimum = 0; // Valor mínimo del ProgressBar
                    progressBar3.Maximum = folio_fact_new.Count; // Valor máximo del ProgressBar
                    progressBar3.Step = 1; // Incremento del ProgressBar por cada iteración
                    Loading3.Visible = true;
                    progressBar3.Visible = true;
                    Loading3.Update();
                    con.Open();
                    string query0 = "SELECT * FROM DOCTOS_VE WHERE (FOLIO LIKE 'PF%' OR FOLIO LIKE 'PE%' OR FOLIO LIKE 'PA%') AND TIPO_DOCTO = 'F' ORDER BY FECHA DESC";
                    FbCommand command0 = new FbCommand(query0, con);
                    // Objeto para leer los datos obtenidos
                    FbDataReader reader0 = command0.ExecuteReader();
                    int conts = 0;
                    //int contar = 0;
                    while (reader0.Read())
                    {
                        string columna0 = reader0.GetString(0);
                        string columna1 = reader0.GetString(4);

                        if (folio_fact_new.Contains(columna1))
                        {
                            int posicion = folio_fact_new.FindIndex(folio_fact_new => folio_fact_new.Contains(columna1));
                            folio_fact_new2[posicion] = columna1;
                            id_new[posicion] = columna0;
                            //MessageBox.Show(columna1);
                            //MessageBox.Show(importe.ToString());
                            FbCommand command = new FbCommand("CALC_TOTALES_DOCTO_VE", con);
                            command.CommandType = CommandType.StoredProcedure;

                            // Parámetro V_DOCTO_VE_ID
                            FbParameter paramV_DOCTO_VE_ID = new FbParameter("V_DOCTO_VE_ID", FbDbType.Integer);
                            paramV_DOCTO_VE_ID.Value = columna0;
                            command.Parameters.Add(paramV_DOCTO_VE_ID);

                            // Parámetro V_SOLO_LECTURA
                            FbParameter paramV_SOLO_LECTURA = new FbParameter("V_SOLO_LECTURA", FbDbType.Char);
                            paramV_SOLO_LECTURA.Size = 1;
                            paramV_SOLO_LECTURA.Value = 'S';
                            command.Parameters.Add(paramV_SOLO_LECTURA);

                            // Parámetro de retorno NUM_RENGLONES
                            FbParameter paramNUM_RENGLONES = new FbParameter();
                            paramNUM_RENGLONES.ParameterName = "NUM_RENGLONES";
                            paramNUM_RENGLONES.Direction = ParameterDirection.Output;
                            paramNUM_RENGLONES.FbDbType = FbDbType.Integer;
                            command.Parameters.Add(paramNUM_RENGLONES);
                            command.ExecuteNonQuery();

                            // Obtener el valor de retorno NUM_RENGLONES
                            int Num_renglones = (int)paramNUM_RENGLONES.Value;
                            renglonesnew2[posicion] = Num_renglones;
                            //  MessageBox.Show(renglonesnew2[posicion].ToString());
                            // MessageBox.Show(posicion.ToString());
                            // MessageBox.Show(folio_fact_new2[posicion]);
                            //sl.SetCellValue("K"+posicion, Num_renglones);
                            progressBar3.Value = conts;
                            progressBar3.Update();
                            for (int i = 8; i < numFilas; ++i)
                            {

                                if (sl.HasCellValue(i, 7))
                                {
                                    Buscar_Folio = sl.GetCellValueAsString(i, 7);
                                    int conta = 9 - Buscar_Folio.Length;
                                    string prefix = Buscar_Folio.Substring(0, 2);
                                    string suffix = Buscar_Folio.Substring(2);
                                    string patch = "";
                                    for (int j = 0; j < conta; j++)
                                    {
                                        patch = patch + "0";
                                    }
                                    Buscar_Folio = prefix + patch + suffix;
                                    if (Buscar_Folio == folio_fact_new2[posicion])
                                    {
                                        sl.SetCellValue("K" + i, renglonesnew2[posicion]);
                                        //MessageBox.Show(Buscar_Folio);
                                        //MessageBox.Show(renglonesnew2[posicion].ToString());
                                        break;
                                    }
                                    //MessageBox.Show(folio_fact_new[l]);

                                }
                            }
                            conts++;
                            if (conts == folio_fact_new.Count)
                                break;
                        }


                    }

                    reader0.Close();
                    int value = 3;
                    for (int j = 0; j < Facturistas.Count; ++j)
                    {

                        string formula = "=SUMIF(A8:A" + numFilas + ",\"" + Facturistas[value - 3] + "\"," + "K8:K" + numFilas + ")";
                        sl.SetCellValue(value, 16, Facturistas[value - 3]);
                        sl.SetCellValue(value, 17, formula);
                        sl.SetRowHeight(value, 15);
                        sl.SetCellValue(value, 18, "=100*(Q" + value + "/R2)");
                        sl.SetCellValue(2, 16, "=SUM(J8:J" + numFilas + ")");
                        sl.SetCellValue(2, 18, "=SUM(Q3:Q" + value + ")");
                        value++;

                    }
                    sl.SetCellValue(2, 17, "=P2*0.0032");
                    sl.SaveAs(selectedFileName);
                    SLDocument doc = new("C:\\Facturistas.xlsx");
                    int k = 1;
                    doc.SetCellValue("D3", "PORCENTAJE DE EMPAQUE");
                    int numFilas2 = doc.GetWorksheetStatistics().NumberOfRows;
                    List<string> Surtidores = new List<string>();
                    while (doc.HasCellValue("A" + k))
                    {
                        Facturistas[k - 1] = (doc.GetCellValueAsString("B" + k));
                        k++;
                    }
                    doc.SaveAs("C:\\Facturistas.xlsx");
                    value = 3;
                    ActualizarExcel(selectedFileName);
                    SLDocument docNew = new(selectedFileName);
                    docNew.SelectWorksheet("Report");
                    DineroGenerado = docNew.GetCellValueAsDecimal("P2");
                    Bono = docNew.GetCellValueAsDecimal("Q2");
                    for (int j = 0; j < Facturistas.Count; ++j)
                    {
                        docNew.SetCellValue(value, 16, Facturistas[value - 3]);
                        string name = docNew.GetCellValueAsString(value, 16);
                        decimal porcentaje = docNew.GetCellValueAsDecimal(value, 18);
                        string subporcentaje = porcentaje.ToString("0.00");
                        // MessageBox.Show(docNew.GetCellValueAsDecimal(value, 18).ToString());
                        datosFacturistas.Add((name, decimal.Parse(subporcentaje)));
                        value++;

                    }

                    docNew.SaveAs(selectedFileName);
                    SLDocument docFinal = new(filePath);
                    docFinal.SelectWorksheet("Final Report");
                    int t = 3;
                    docFinal.SetCellValue("B3", "PORCENTAJE DE FACTURACION");
                    docFinal.SetCellValue("E3", "MONTO DE FACTURACION");
                    docFinal.SetCellValue("F3", "MONTO DE SURTIDO");
                    docFinal.SetCellValue("G3", "MONTO DE EMPAQUE");
                    docFinal.SetCellValue("H3", "ACUMULADO");
                    docFinal.SetCellValue("A1", "BONO CORRESPONDIENTE");
                    docFinal.SetCellValue("B1", "% FACTURACIÓN");
                    docFinal.SetCellValue("C1", "% SURTIDO");
                    docFinal.SetCellValue("D1", "% EMPAQUE");
                    docFinal.SetCellValue("A2", Bono);
                    docFinal.SetCellStyle("A1", style2);
                    docFinal.SetCellStyle("B1", style2);
                    docFinal.SetCellStyle("C1", style2);
                    docFinal.SetCellStyle("D1", style2);
                    docFinal.SetCellStyle("A2", style3);
                    docFinal.SetCellStyle("B2", style5);
                    docFinal.SetCellStyle("C2", style5);
                    docFinal.SetCellStyle("D2", style5);
                    docFinal.SetCellStyle("A3", style2);
                    docFinal.SetCellStyle("B3", style2);
                    docFinal.SetCellStyle("C3", style2);
                    docFinal.SetCellStyle("D3", style2);
                    docFinal.SetCellStyle("E3", style2);
                    docFinal.SetCellStyle("F3", style2);
                    docFinal.SetCellStyle("A3", style2);
                    docFinal.SetCellStyle("G3", style2);
                    docFinal.SetCellStyle("H3", style2);
                    int numFilas3 = docFinal.GetWorksheetStatistics().NumberOfRows;
                    List<string> SurtidoresAct = new List<string>();
                    while (docFinal.HasCellValue("A" + t))
                    {
                        SurtidoresAct.Add(docFinal.GetCellValueAsString("A" + t));
                        t++;
                    }

                    foreach (var tupla2 in datosFacturistas)
                    {

                        if (SurtidoresAct.Contains(tupla2.Item1))
                        {
                            int posicion = SurtidoresAct.FindIndex(SurtidoresAct => SurtidoresAct.Contains(tupla2.Item1));
                            //   MessageBox.Show(posicion.ToString());
                            // MessageBox.Show(Surtidores[posicion]);
                            // doc.SetCellValue("A" + i, tupla.Item1);
                            posicion = posicion + 3;
                            docFinal.SetCellValue("B" + posicion, tupla2.Item2);
                        }
                        else
                        {
                            numFilas3 = numFilas3 + 3;
                            docFinal.SetCellValue("A" + numFilas3, tupla2.Item1);
                            docFinal.SetCellValue("B" + numFilas3, tupla2.Item2);
                            numFilas3 = docFinal.GetWorksheetStatistics().NumberOfRows;
                        }
                    }
                    int final = numFilas3 + 1;
                    int final2 = numFilas3 + 2;
                    docFinal.SetCellValue("B" + final, "CORRESPONDE");
                    docFinal.SetCellValue("C" + final, "CORRESPONDE");
                    docFinal.SetCellValue("D" + final, "CORRESPONDE");
                    docFinal.SetCellValue("B" + final2, "=ROUND(SUM(B2*A2),2)");
                    docFinal.SetCellValue("C" + final2, "=ROUND(SUM(C2*A2),2)");
                    docFinal.SetCellValue("D" + final2, "=ROUND(SUM(D2*A2),2)");
                    docFinal.SetCellStyle("B" + final2, style3);
                    docFinal.SetCellStyle("C" + final2, style3);
                    docFinal.SetCellStyle("D" + final2, style3);
                    docFinal.SetCellStyle("B" + final, style2);
                    docFinal.SetCellStyle("C" + final, style2);
                    docFinal.SetCellStyle("D" + final, style2);

                    for (int i = 4; i < numFilas3+1; i++) {
                        docFinal.SetCellValue("E" + i, "=ROUND((B" + i + "%*B" + final2 + "),2)");
                        docFinal.SetCellValue("F" + i, "=ROUND((C" + i + "%*C" + final2 + "),2)");
                        docFinal.SetCellValue("G" + i, "=ROUND((D" + i + "%*D" + final2 + "),2)");
                        docFinal.SetCellValue("H" + i, "=ROUND((E" + i + "+F" + i + "+G" + i+"),2)");
                        docFinal.SetCellStyle("H"+i, style3);
                        docFinal.SetCellStyle("E" + i, style6);
                        docFinal.SetCellStyle("F" + i, style6);
                        docFinal.SetCellStyle("G" + i, style6);
                        docFinal.SetCellStyle("B" + i, style);
                        docFinal.SetCellStyle("C" + i, style);
                        docFinal.SetCellStyle("D" + i, style);
                        docFinal.SetCellStyle("A" + i, style2);

                    }
                    docFinal.SaveAs(filePath);
                    progressBar3.Visible = false;
                    Loading3.Visible = false;
                    MessageBox.Show("Cargado con Éxito");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error de conexión, vuelve a intentar o contacta a 06", "¡Espera!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(ex.ToString());

                    //sl.SaveAs(openFileDialog1.FileName);
                }
            }

        }
        private void Btn_Empaque_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;";

            // Show the dialog and check if the user clicked the OK button
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog1.FileName;

                progressBar2.Minimum = 0; // Valor mínimo del ProgressBar
                progressBar2.Step = 1; // Incremefolio_fact_new.Countnto del ProgressBar por cada iteración
                Loading2.Visible = true;
                progressBar2.Visible = true;
                Loading2.Update();

                using (SLDocument sl = new SLDocument(selectedFileName))
                {

                    sl.SelectWorksheet("Promedio");
                    int numFilas = sl.GetWorksheetStatistics().NumberOfRows;
                    // Definimos el rango de la columna A que queremos ordenar (ejemplo: de la fila 2 a la 11)
                    int filaInicio = 2;
                    int filaFin = numFilas;

                    // Creamos una lista para almacenar los datos de las filas y la columna A

                    // Leer los datos de las filas y la columna A
                    for (int fila = filaInicio; fila <= filaFin; fila++)
                    {
                        string valorColumnaA = sl.GetCellValueAsString(fila, 1);
                        decimal valorColumnaG = sl.GetCellValueAsDecimal(fila, 7);
                        string subporcentaje = valorColumnaG.ToString("0.00");
                        datosFilas.Add((valorColumnaA, decimal.Parse(subporcentaje)));
                    }
                    datosFilas.Sort();
                    progressBar2.Maximum = datosFilas.Count; // Valor máximo del ProgressBar
                    // Guardar los cambios en el archivo ordenado
                    sl.SaveAs(selectedFileName);

                }
                SLDocument doc = new(filePath);
                doc.SelectWorksheet("Final Report");
                int i = 3;
                doc.SetCellValue("D3", "PORCENTAJE DE EMPAQUE");
                int numFilas2 = doc.GetWorksheetStatistics().NumberOfRows;
                List<string> Surtidores = new List<string>();
                while (doc.HasCellValue("A" + i))
                {
                    Surtidores.Add(doc.GetCellValueAsString("A" + i));
                    i++;
                }
                int contar = 0;
                foreach (var tupla in datosFilas)
                {
                    progressBar2.Value = contar++;
                    progressBar2.Update();
                    numFilas2 = doc.GetWorksheetStatistics().NumberOfRows;
                    if (Surtidores.Contains(tupla.Item1))
                    {
                        int posicion = Surtidores.FindIndex(Surtidores => Surtidores.Contains(tupla.Item1));
                        //   MessageBox.Show(posicion.ToString());
                        // MessageBox.Show(Surtidores[posicion]);
                        // doc.SetCellValue("A" + i, tupla.Item1);
                        posicion = posicion + 3;
                        doc.SetCellValue("D" + posicion, tupla.Item2);
                    }
                    else
                    {
                        numFilas2 = numFilas2 + 3;
                        doc.SetCellValue("A" + numFilas2, tupla.Item1);
                        doc.SetCellValue("D" + numFilas2, tupla.Item2);
                    }
                }
                progressBar2.Visible = false;
                Loading2.Visible = false;
                doc.SaveAs(filePath);
                MessageBox.Show("Cargado con Éxito");
            }


        }
        public void ActualizarExcel(string filename)
        {

            using (var package = new ExcelPackage(new System.IO.FileInfo(filename)))
            {
                // Recorrer todas las hojas de cálculo y recalcular las fórmulas
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    worksheet.Calculate();
                }
                package.Save();

            }

        }
        
    }

}
