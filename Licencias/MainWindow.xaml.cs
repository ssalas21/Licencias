using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Licencias
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int max = Convert.ToInt32(TxtMax.Text);
            int numeroArchivo = Convert.ToInt32(TxtNumArchivo.Text);
            string numInforme = TxtNumInforme.Text;
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\LPNV" + numInforme + ".txt"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("PUCHUNCAVI");

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Licencias\\Licencias" + numeroArchivo + ".xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar fila
                for (rCnt = 2; rCnt <= max; rCnt++)
                {
                    string linea = "L";
                    string ingreso = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                    if (ingreso.Equals("Duplicado")) linea = linea + "2CA";
                    else linea = linea + "1CA";
                    string folio = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    while (folio.Length < 10)
                    {
                        folio = "0" + folio;
                    }
                    linea = linea + folio;
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 3] as Excel.Range).Value2));
                    linea = linea + fecha.Year.ToString();
                    if (fecha.Month < 10) linea = linea + "0" + fecha.Month.ToString();
                    else linea = linea + fecha.Month.ToString();
                    if (fecha.Day < 10) linea = linea + "0" + fecha.Day.ToString();
                    else linea = linea + fecha.Day.ToString();
                    string clase = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    string clase2 = "";
                    if (clase.IndexOf("0,") != -1) clase2 = clase2 + "B ";
                    if (clase.IndexOf("1,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("2,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("3,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("4,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("5,") != -1) clase2 = clase2 + "A3";
                    if (clase.IndexOf("6,") != -1) clase2 = clase2 + "A4";
                    if (clase.IndexOf("7,") != -1) clase2 = clase2 + "A5";
                    if (clase.IndexOf("8,") != -1) clase2 = clase2 + "C ";
                    if (clase.IndexOf("9,") != -1) clase2 = clase2 + "D ";
                    if (clase.IndexOf("10,") != -1) clase2 = clase2 + "E ";
                    if (clase.IndexOf("11,") != -1) clase2 = clase2 + "F ";
                    int aux2 = clase2.Length;
                    if (clase2.Length < 20)
                    {
                        for (int k = aux2; k < 20; k++)
                        {
                            clase2 = clase2 + " ";
                        }
                    }
                    else
                    {
                        clase2 = clase2.Substring(0, 20);
                    }
                    linea = linea + clase2;
                    string rut = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    linea = linea + rut.Substring(0, rut.IndexOf("-"));
                    linea = linea + rut.Substring(rut.IndexOf("-") + 1, 1);
                    string apellido1 = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                    if (apellido1.Length < 30)
                    {
                        int aux = apellido1.Length;
                        for (int k = aux; k < 30; k++)
                        {
                            apellido1 = apellido1 + " ";
                        }
                    }
                    linea = linea + apellido1;
                    try
                    {
                        string apellido2 = ((range.Cells[rCnt, 7] as Excel.Range).Value2).ToString();
                        if (apellido2.Length < 30)
                        {
                            int aux = apellido2.Length;
                            for (int k = aux; k < 30; k++)
                            {
                                apellido2 = apellido2 + " ";
                            }
                        }
                        linea = linea + apellido2;
                    }
                    catch (Exception)
                    {
                        linea = linea + "                              ";

                    }
                    string nombres = ((range.Cells[rCnt, 8] as Excel.Range).Value2).ToString();
                    if (nombres.Length < 40)
                    {
                        int aux = nombres.Length;
                        for (int k = aux; k < 40; k++)
                        {
                            nombres = nombres + " ";
                        }
                    }
                    linea = linea + nombres;

                    try
                    {
                        string rutescuela = ((range.Cells[rCnt, 9] as Excel.Range).Value2).ToString();
                        linea = linea + rutescuela.Substring(0, rutescuela.IndexOf("-"));
                        linea = linea + rutescuela.Substring(rutescuela.IndexOf("-") + 1, 1);
                        string nombreescuela = ((range.Cells[rCnt, 14] as Excel.Range).Value2).ToString();
                        if (nombreescuela.Length < 40)
                        {
                            int aux = nombreescuela.Length;
                            for (int k = aux; k < 40; k++)
                            {
                                nombreescuela = nombreescuela + " ";
                            }
                        }
                        linea = linea + nombreescuela;
                        DateTime fechaaprobacion = DateTime.FromOADate(((range.Cells[rCnt, 10] as Excel.Range).Value2));
                        linea = linea + fechaaprobacion.Year.ToString();
                        if (fechaaprobacion.Month < 10) linea = linea + "0" + fechaaprobacion.Month.ToString();
                        else linea = linea + fechaaprobacion.Month.ToString();
                        if (fechaaprobacion.Day < 10) linea = linea + "0" + fechaaprobacion.Day.ToString();
                        else linea = linea + fechaaprobacion.Day.ToString();
                    }
                    catch (Exception)
                    {
                        linea = linea + "00000000                                         00000000";
                    }

                    file.WriteLine(linea);
                    linea = "";
                    linea = "D";
                    string comuna = ((range.Cells[rCnt, 19] as Excel.Range).Value2).ToString();
                    if (comuna.Length < 55)
                    {
                        int aux = comuna.Length;
                        for (int k = aux; k < 55; k++)
                        {
                            comuna = comuna + " ";
                        }
                    }
                    linea = linea + comuna;
                    string direccion = ((range.Cells[rCnt, 15] as Excel.Range).Value2).ToString();
                    if (direccion.Length < 45)
                    {
                        int aux = direccion.Length;
                        for (int k = aux; k < 45; k++)
                        {
                            direccion = direccion + " ";
                        }
                    }
                    linea = linea + direccion;
                    string direccion2 = ((range.Cells[rCnt, 16] as Excel.Range).Value2).ToString();
                    if (direccion2.Length < 9)
                    {
                        int aux = direccion2.Length;
                        for (int k = aux; k < 9; k++)
                        {
                            direccion2 = direccion2 + " ";
                        }
                    }

                    linea = linea + direccion2;
                    try
                    {
                        string letra = ((range.Cells[rCnt, 17] as Excel.Range).Value2).ToString();
                        if (letra.Length < 3)
                        {
                            int aux = letra.Length;
                            for (int k = aux; k < 3; k++)
                            {
                                letra = letra + " ";
                            }
                        }
                        linea = linea + letra;

                    }
                    catch (Exception)
                    {

                        linea = linea + "   ";
                    }

                    try
                    {
                        string resto = ((range.Cells[rCnt, 18] as Excel.Range).Value2).ToString();
                        if (resto.Length < 3)
                        {
                            int aux = resto.Length;
                            for (int k = aux; k < 3; k++)
                            {
                                resto = resto + " ";
                            }
                        }
                        linea = linea + resto;

                    }
                    catch (Exception)
                    {

                        linea = linea + "                                             ";
                    }

                    file.WriteLine(linea);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            int max = Convert.ToInt32(TxtMax.Text);
            int numeroArchivo = Convert.ToInt32(TxtNumArchivo.Text);
            string numInforme = TxtNumInforme.Text;
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\RPNV" + numInforme + ".txt"); // Abrir el txt
            // Cabeceras del html
            file.WriteLine("PUCHUNCAVI");

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Licencias\\conduce" + numeroArchivo + ".xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 2; rCnt <= max; rCnt++)
                {
                    string linea = "L";
                    string ingreso = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                    if (ingreso.Equals("Duplicado")) linea = linea + "2CA";
                    else linea = linea + "1CA";
                    string folio = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    while (folio.Length < 10)
                    {
                        folio = "0" + folio;
                    }
                    linea = linea + folio;
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 3] as Excel.Range).Value2));
                    linea = linea + fecha.Year.ToString();
                    if (fecha.Month < 10) linea = linea + "0" + fecha.Month.ToString();
                    else linea = linea + fecha.Month.ToString();
                    if (fecha.Day < 10) linea = linea + "0" + fecha.Day.ToString();
                    else linea = linea + fecha.Day.ToString();
                    string clase = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    string clase2 = "";
                    int aux2 = 0;
                    if (clase.IndexOf("0,") != -1) clase2 = clase2 + "B ";
                    if (clase.IndexOf("1,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("2,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("3,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("4,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("5,") != -1) clase2 = clase2 + "A3";
                    if (clase.IndexOf("6,") != -1) clase2 = clase2 + "A4";
                    if (clase.IndexOf("7,") != -1) clase2 = clase2 + "A5";
                    if (clase.IndexOf("8,") != -1) clase2 = clase2 + "C ";
                    if (clase.IndexOf("9,") != -1) clase2 = clase2 + "D ";
                    if (clase.IndexOf("10,") != -1) clase2 = clase2 + "E ";
                    if (clase.IndexOf("11,") != -1) clase2 = clase2 + "F ";
                    aux2 = clase2.Length;
                    if (clase2.Length < 20)
                    {
                        for (int k = aux2; k < 20; k++)
                        {
                            clase2 = clase2 + " ";
                        }
                    }
                    else
                    {
                        clase2 = clase2.Substring(0, 20);
                    }
                    linea = linea + clase2;
                    string rut = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    linea = linea + rut.Substring(0, rut.IndexOf("-"));
                    linea = linea + rut.Substring(rut.IndexOf("-") + 1, 1);
                    string apellido1 = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                    if (apellido1.Length < 30)
                    {
                        int aux = apellido1.Length;
                        for (int k = aux; k < 30; k++)
                        {
                            apellido1 = apellido1 + " ";
                        }
                    }
                    linea = linea + apellido1;
                    try
                    {
                        string apellido2 = ((range.Cells[rCnt, 7] as Excel.Range).Value2).ToString();
                        if (apellido2.Length < 30)
                        {
                            int aux = apellido2.Length;
                            for (int k = aux; k < 30; k++)
                            {
                                apellido2 = apellido2 + " ";
                            }
                        }
                        linea = linea + apellido2;
                    }
                    catch (Exception)
                    {
                        linea = linea + "                              ";

                    }
                    string nombres = ((range.Cells[rCnt, 8] as Excel.Range).Value2).ToString();
                    if (nombres.Length < 40)
                    {
                        int aux = nombres.Length;
                        for (int k = aux; k < 40; k++)
                        {
                            nombres = nombres + " ";
                        }
                    }
                    linea = linea + nombres;

                    DateTime fechaexpiracion = DateTime.FromOADate(((range.Cells[rCnt, 13] as Excel.Range).Value2));
                    linea = linea + fechaexpiracion.Year.ToString();
                    if (fechaexpiracion.Month < 10) linea = linea + "0" + fechaexpiracion.Month.ToString();
                    else linea = linea + fechaexpiracion.Month.ToString();
                    if (fechaexpiracion.Day < 10) linea = linea + "0" + fechaexpiracion.Day.ToString();
                    else linea = linea + fechaexpiracion.Day.ToString();

                    string restriccion = ((range.Cells[rCnt, 11] as Excel.Range).Value2).ToString();
                    string restriccion2 = "";
                    int aux3 = 0;
                    if (restriccion.IndexOf("SIN") != -1) restriccion2 = restriccion2 + "00";
                    if (restriccion.IndexOf("1,") != -1) restriccion2 = restriccion2 + "01";
                    if (restriccion.IndexOf("2,") != -1) restriccion2 = restriccion2 + "02";
                    if (restriccion.IndexOf("3,") != -1) restriccion2 = restriccion2 + "03";
                    if (restriccion.IndexOf("4,") != -1) restriccion2 = restriccion2 + "04";
                    if (restriccion.IndexOf("5,") != -1) restriccion2 = restriccion2 + "05";
                    if (restriccion.IndexOf("6,") != -1) restriccion2 = restriccion2 + "06";
                    if (restriccion.IndexOf("7,") != -1) restriccion2 = restriccion2 + "07";
                    if (restriccion.IndexOf("8,") != -1) restriccion2 = restriccion2 + "08";
                    if (restriccion.IndexOf("9,") != -1) restriccion2 = restriccion2 + "09";
                    if (restriccion.IndexOf("10,") != -1) restriccion2 = restriccion2 + "10";
                    aux3 = restriccion2.Length;
                    if (restriccion2.Length < 16)
                    {
                        for (int k = aux3; k < 16; k++)
                        {
                            restriccion2 = restriccion2 + "0";
                        }
                    }
                    else
                    {
                        restriccion2 = restriccion2.Substring(0, 16);
                    }
                    linea = linea + restriccion2;
                    try
                    {
                        string rutescuela = ((range.Cells[rCnt, 9] as Excel.Range).Value2).ToString();
                        linea = linea + rutescuela.Substring(0, rutescuela.IndexOf("-"));
                        linea = linea + rutescuela.Substring(rutescuela.IndexOf("-") + 1, 1);
                        string nombreescuela = ((range.Cells[rCnt, 14] as Excel.Range).Value2).ToString();
                        if (nombreescuela.Length < 40)
                        {
                            int aux = nombreescuela.Length;
                            for (int k = aux; k < 40; k++)
                            {
                                nombreescuela = nombreescuela + " ";
                            }
                        }
                        linea = linea + nombreescuela;
                        DateTime fechaaprobacion = DateTime.FromOADate(((range.Cells[rCnt, 10] as Excel.Range).Value2));
                        linea = linea + fechaaprobacion.Year.ToString();
                        if (fechaaprobacion.Month < 10) linea = linea + "0" + fechaaprobacion.Month.ToString();
                        else linea = linea + fechaaprobacion.Month.ToString();
                        if (fechaaprobacion.Day < 10) linea = linea + "0" + fechaaprobacion.Day.ToString();
                        else linea = linea + fechaaprobacion.Day.ToString();
                    }
                    catch (Exception)
                    {
                        linea = linea + "00000000                                         00000000";
                    }

                    file.WriteLine(linea);
                    linea = "";
                    linea = "D";
                    string comuna = ((range.Cells[rCnt, 19] as Excel.Range).Value2).ToString();
                    if (comuna.Length < 55)
                    {
                        int aux = comuna.Length;
                        for (int k = aux; k < 55; k++)
                        {
                            comuna = comuna + " ";
                        }
                    }
                    linea = linea + comuna;
                    string direccion = ((range.Cells[rCnt, 15] as Excel.Range).Value2).ToString();
                    if (direccion.Length < 45)
                    {
                        int aux = direccion.Length;
                        for (int k = aux; k < 45; k++)
                        {
                            direccion = direccion + " ";
                        }
                    }
                    linea = linea + direccion;
                    string direccion2 = ((range.Cells[rCnt, 16] as Excel.Range).Value2).ToString();
                    if (direccion2.Length < 9)
                    {
                        int aux = direccion2.Length;
                        for (int k = aux; k < 9; k++)
                        {
                            direccion2 = direccion2 + " ";
                        }
                    }

                    linea = linea + direccion2;
                    try
                    {
                        string letra = ((range.Cells[rCnt, 17] as Excel.Range).Value2).ToString();
                        if (letra.Length < 3)
                        {
                            int aux = letra.Length;
                            for (int k = aux; k < 3; k++)
                            {
                                letra = letra + " ";
                            }
                        }
                        linea = linea + letra;

                    }
                    catch (Exception)
                    {

                        linea = linea + "   ";
                    }

                    try
                    {
                        string resto = ((range.Cells[rCnt, 18] as Excel.Range).Value2).ToString();
                        if (resto.Length < 3)
                        {
                            int aux = resto.Length;
                            for (int k = aux; k < 3; k++)
                            {
                                resto = resto + " ";
                            }
                        }
                        linea = linea + resto;

                    }
                    catch (Exception)
                    {

                        linea = linea + "                                             ";
                    }

                    file.WriteLine(linea);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            int max = Convert.ToInt32(TxtMax.Text);
            int numeroArchivo = Convert.ToInt32(TxtNumArchivo.Text);
            string numInforme = TxtNumInforme.Text;
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\DPNV" + numInforme + ".txt"); // Abrir el txt
            // Cabeceras del html

            file.WriteLine("PUCHUNCAVI");
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Licencias\\denegaciones" + numeroArchivo + ".xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);



            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 2; rCnt <= max; rCnt++)
                {

                    //string linea = "L";
                    string linea = "";
                    //string ingreso = ((range.Cells[rCnt, 1] as Excel.Range).Value2).ToString();
                    //if (ingreso.Equals("Duplicado")) linea = linea + "2CA";
                    //else linea = linea + "1CA";
                    //string folio = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    //while (folio.Length < 10)
                    //{
                    //    folio = "0" + folio;
                    //}
                    //linea = linea + folio;
                    DateTime fecha = DateTime.FromOADate(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    linea = linea + fecha.Year.ToString();
                    if (fecha.Month < 10) linea = linea + "0" + fecha.Month.ToString();
                    else linea = linea + fecha.Month.ToString();
                    if (fecha.Day < 10) linea = linea + "0" + fecha.Day.ToString();
                    else linea = linea + fecha.Day.ToString();
                    string motivo = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    linea = linea + motivo;
                    string duracion = ((range.Cells[rCnt, 3] as Excel.Range).Value2).ToString();
                    linea = linea + duracion;
                    string clase = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    string clase2 = "";
                    int aux2 = 0;
                    if (clase.IndexOf("0,") != -1) clase2 = clase2 + "B ";
                    if (clase.IndexOf("1,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("2,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("3,") != -1) clase2 = clase2 + "A1";
                    if (clase.IndexOf("4,") != -1) clase2 = clase2 + "A2";
                    if (clase.IndexOf("5,") != -1) clase2 = clase2 + "A3";
                    if (clase.IndexOf("6,") != -1) clase2 = clase2 + "A4";
                    if (clase.IndexOf("7,") != -1) clase2 = clase2 + "A5";
                    if (clase.IndexOf("8,") != -1) clase2 = clase2 + "C ";
                    if (clase.IndexOf("9,") != -1) clase2 = clase2 + "D ";
                    if (clase.IndexOf("10,") != -1) clase2 = clase2 + "E ";
                    if (clase.IndexOf("11,") != -1) clase2 = clase2 + "F ";
                    aux2 = clase2.Length;
                    if (clase2.Length < 20)
                    {
                        for (int k = aux2; k < 20; k++)
                        {
                            clase2 = clase2 + " ";
                        }
                    }
                    else
                    {
                        clase2 = clase2.Substring(0, 20);
                    }
                    linea = linea + clase2;
                    string rut = ((range.Cells[rCnt, 5] as Excel.Range).Value2).ToString();
                    linea = linea + rut.Substring(0, rut.IndexOf("-"));
                    linea = linea + rut.Substring(rut.IndexOf("-") + 1, 1);
                    string apellido1 = ((range.Cells[rCnt, 6] as Excel.Range).Value2).ToString();
                    if (apellido1.Length < 20)
                    {
                        int aux = apellido1.Length;
                        for (int k = aux; k < 20; k++)
                        {
                            apellido1 = apellido1 + " ";
                        }
                    }
                    linea = linea + apellido1;
                    try
                    {
                        string apellido2 = ((range.Cells[rCnt, 7] as Excel.Range).Value2).ToString();
                        if (apellido2.Length < 20)
                        {
                            int aux = apellido2.Length;
                            for (int k = aux; k < 20; k++)
                            {
                                apellido2 = apellido2 + " ";
                            }
                        }
                        linea = linea + apellido2;
                    }
                    catch (Exception)
                    {
                        linea = linea + "                              ";

                    }
                    string nombres = ((range.Cells[rCnt, 8] as Excel.Range).Value2).ToString();
                    if (nombres.Length < 30)
                    {
                        int aux = nombres.Length;
                        for (int k = aux; k < 40; k++)
                        {
                            nombres = nombres + " ";
                        }
                    }
                    linea = linea + nombres;

                    //DateTime fechaexpiracion = DateTime.FromOADate(((range.Cells[rCnt, 13] as Excel.Range).Value2));
                    //linea = linea + fechaexpiracion.Year.ToString();
                    //if (fechaexpiracion.Month < 10) linea = linea + "0" + fechaexpiracion.Month.ToString();
                    //else linea = linea + fechaexpiracion.Month.ToString();
                    //if (fechaexpiracion.Day < 10) linea = linea + "0" + fechaexpiracion.Day.ToString();
                    //else linea = linea + fechaexpiracion.Day.ToString();

                    //string restriccion = ((range.Cells[rCnt, 11] as Excel.Range).Value2).ToString();
                    //string restriccion2 = "";
                    //int aux3 = 0;
                    //if (restriccion.IndexOf("SIN") != -1) restriccion2 = restriccion2 + "00";
                    //if (restriccion.IndexOf("1,") != -1) restriccion2 = restriccion2 + "01";
                    //if (restriccion.IndexOf("2,") != -1) restriccion2 = restriccion2 + "02";
                    //if (restriccion.IndexOf("3,") != -1) restriccion2 = restriccion2 + "03";
                    //if (restriccion.IndexOf("4,") != -1) restriccion2 = restriccion2 + "04";
                    //if (restriccion.IndexOf("5,") != -1) restriccion2 = restriccion2 + "05";
                    //if (restriccion.IndexOf("6,") != -1) restriccion2 = restriccion2 + "06";
                    //if (restriccion.IndexOf("7,") != -1) restriccion2 = restriccion2 + "07";
                    //if (restriccion.IndexOf("8,") != -1) restriccion2 = restriccion2 + "08";
                    //if (restriccion.IndexOf("9,") != -1) restriccion2 = restriccion2 + "09";
                    //if (restriccion.IndexOf("10,") != -1) restriccion2 = restriccion2 + "10";
                    //aux3 = restriccion2.Length;
                    //if (restriccion2.Length < 16)
                    //{
                    //    for (int k = aux3; k < 16; k++)
                    //    {
                    //        restriccion2 = restriccion2 + "0";
                    //    }
                    //}
                    //else
                    //{
                    //    restriccion2 = restriccion2.Substring(0, 16);
                    //}
                    //linea = linea + restriccion2;
                    //try
                    //{
                    //    string rutescuela = ((range.Cells[rCnt, 9] as Excel.Range).Value2).ToString();
                    //    linea = linea + rutescuela.Substring(0, rutescuela.IndexOf("-"));
                    //    linea = linea + rutescuela.Substring(rutescuela.IndexOf("-") + 1, 1);
                    //    string nombreescuela = ((range.Cells[rCnt, 14] as Excel.Range).Value2).ToString();
                    //    if (nombreescuela.Length < 40)
                    //    {
                    //        int aux = nombreescuela.Length;
                    //        for (int k = aux; k < 40; k++)
                    //        {
                    //            nombreescuela = nombreescuela + " ";
                    //        }
                    //    }
                    //    linea = linea + nombreescuela;
                    //    DateTime fechaaprobacion = DateTime.FromOADate(((range.Cells[rCnt, 10] as Excel.Range).Value2));
                    //    linea = linea + fechaaprobacion.Year.ToString();
                    //    if (fechaaprobacion.Month < 10) linea = linea + "0" + fechaaprobacion.Month.ToString();
                    //    else linea = linea + fechaaprobacion.Month.ToString();
                    //    if (fechaaprobacion.Day < 10) linea = linea + "0" + fechaaprobacion.Day.ToString();
                    //    else linea = linea + fechaaprobacion.Day.ToString();
                    //}
                    //catch (Exception)
                    //{
                    //    linea = linea + "00000000                                         00000000";
                    //}

                    //file.WriteLine(linea);
                    //linea = "";
                    //linea = "D";
                    //string comuna = ((range.Cells[rCnt, 19] as Excel.Range).Value2).ToString();
                    //if (comuna.Length < 55)
                    //{
                    //    int aux = comuna.Length;
                    //    for (int k = aux; k < 55; k++)
                    //    {
                    //        comuna = comuna + " ";
                    //    }
                    //}
                    //linea = linea + comuna;
                    //string direccion = ((range.Cells[rCnt, 15] as Excel.Range).Value2).ToString();
                    //if (direccion.Length < 45)
                    //{
                    //    int aux = direccion.Length;
                    //    for (int k = aux; k < 45; k++)
                    //    {
                    //        direccion = direccion + " ";
                    //    }
                    //}
                    //linea = linea + direccion;
                    //string direccion2 = ((range.Cells[rCnt, 16] as Excel.Range).Value2).ToString();
                    //if (direccion2.Length < 9)
                    //{
                    //    int aux = direccion2.Length;
                    //    for (int k = aux; k < 9; k++)
                    //    {
                    //        direccion2 = direccion2 + " ";
                    //    }
                    //}

                    //linea = linea + direccion2;
                    //try
                    //{
                    //    string letra = ((range.Cells[rCnt, 17] as Excel.Range).Value2).ToString();
                    //    if (letra.Length < 3)
                    //    {
                    //        int aux = letra.Length;
                    //        for (int k = aux; k < 3; k++)
                    //        {
                    //            letra = letra + " ";
                    //        }
                    //    }
                    //    linea = linea + letra;

                    //}
                    //catch (Exception)
                    //{

                    //    linea = linea + "   ";
                    //}

                    //try
                    //{
                    //    string resto = ((range.Cells[rCnt, 18] as Excel.Range).Value2).ToString();
                    //    if (resto.Length < 3)
                    //    {
                    //        int aux = resto.Length;
                    //        for (int k = aux; k < 3; k++)
                    //        {
                    //            resto = resto + " ";
                    //        }
                    //    }
                    //    linea = linea + resto;

                    //}
                    //catch (Exception)
                    //{

                    //    linea = linea + "                                             ";
                    //}

                    file.WriteLine(linea);
                }

                Marshal.ReleaseComObject(xlWorkSheet2);

            }

            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);


            //Cerrar archivo
            file.Close();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\Estadistica\\Estadistica.txt");
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;

            int rCnt;

            int rw = 0;
            int cl = 0;

            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Estadistica\\Datos.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            List<Licencia> listado = new List<Licencia>();

            List<Licencia> listado2 = new List<Licencia>()
                {
                    new Licencia(){ Tipo="B", Cantidad=0},
                    new Licencia(){ Tipo="A1", Cantidad=0},
                    new Licencia(){ Tipo="A2", Cantidad=0},
                    new Licencia(){ Tipo="A3", Cantidad=0},
                    new Licencia(){ Tipo="A4", Cantidad=0},
                    new Licencia(){ Tipo="C", Cantidad=0},
                    new Licencia(){ Tipo="D", Cantidad=0},
                    new Licencia(){ Tipo="E", Cantidad=0},
                    new Licencia(){ Tipo="F", Cantidad=0}
                }
                ;
            for (int i = 1; i <= 1; i++)
            {

                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);

                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                //Llenar filas
                for (rCnt = 2; rCnt <= 6272; rCnt++)
                {
                    int bandera = 0;
                    string licencia = ((range.Cells[rCnt, 4] as Excel.Range).Value2).ToString();
                    foreach (Licencia item in listado)
                    {
                        if (licencia.Equals(item.Tipo))
                        {
                            item.Cantidad = item.Cantidad + 1;
                            bandera = 1;
                        }
                    }
                    foreach (Licencia item in listado2)
                    {
                        if (licencia.IndexOf(item.Tipo) != -1)
                        {
                            item.Cantidad = item.Cantidad + 1;
                        }
                    }
                    if (bandera == 0) listado.Add(new Licencia() { Tipo = licencia, Cantidad = 1 });


                }

                file.WriteLine("La cantidad de licencias entregadas es de 6271");
                foreach (Licencia item in listado2)
                {
                    file.WriteLine(string.Format("Documentos de licencias que se han entregado con clase " + item.Tipo + " incluida = " + item.Cantidad));
                }
                file.WriteLine("");
                file.WriteLine("Las combinaciones de licencias son las siguientes");
                foreach (Licencia item in listado)
                {
                    file.WriteLine(string.Format("Documentos de licencias que se han entregado con las clases " + item.Tipo + " = " + item.Cantidad));
                }
                Marshal.ReleaseComObject(xlWorkSheet2);
                int cantidad = 0;
                foreach (Licencia item in listado)
                {
                    cantidad = cantidad + item.Cantidad;
                }
                file.WriteLine(string.Format("La cantidad es " + cantidad));
            }
            file.Close();
            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();


            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);
        }

        private void BtnIne_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }

            // Cargar el excel conduce

            Excel.Application xlAppConduce;
            Excel.Workbook xlWorkBookConduce;
            Excel.Worksheet xlWorkSheetConduce;
            Excel.Range rangeConduce;

            xlAppConduce = new Excel.Application();
            xlWorkBookConduce = xlAppConduce.Workbooks.Open("D:\\INE\\Licencias\\Ine\\2019\\base.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            xlWorkSheetConduce = (Excel.Worksheet)xlWorkBookConduce.Worksheets.get_Item(1);
            rangeConduce = xlWorkSheetConduce.UsedRange;


            //Crear archivo
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "MES";
            xlWorkSheet.Cells[1, 2] = "Clase de Licencia";
            xlWorkSheet.Cells[1, 3] = "Tipo de Tramite";
            xlWorkSheet.Cells[1, 4] = "Genero";
            xlWorkSheet.Cells[1, 5] = "Fecha Nacimiento";
            xlWorkSheet.Cells[1, 6] = "Nombre";
            xlWorkSheet.Cells[1, 7] = "Rut";

            int rowExcel = 2;

            for (int i = 2; i <= 1184; i++)
            {
                int peticion = Convert.ToInt32((rangeConduce.Cells[i, 4] as Excel.Range).Value2);
                //Peticion 1 = Nueva licencia
                if (peticion == 1)
                {
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {
                        xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                        xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                        xlWorkSheet.Cells[rowExcel, 3] = "NUEVA";
                        string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                        if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                        else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                        xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                        xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                        xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                        xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                        rowExcel++;
                    }
                }
                //Peticion 2 = Renovacion
                if (peticion == 2)
                {
                    //string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    //string[] stringSeparators = new string[] { "," };
                    //string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    //string licencia2 = (rangeConduce.Cells[i, 3] as Excel.Range).Value2.ToString();
                    //string[] resultado2 = licencia2.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    //for (int j = 0; j < resultado.Length; j++)
                    //{
                    //    bool flag = true;
                    //    for (int k = 0; k < resultado2.Length; k++)
                    //    {
                    //        if (resultado[j].Trim() == resultado2[k].Trim()) flag = false;
                    //    }
                    //    if (flag == true)
                    //    {
                    //        xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                    //        xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                    //        xlWorkSheet.Cells[rowExcel, 3] = "RENOVACIÓN";
                    //        string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                    //        if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                    //        else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                    //        xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                    //        xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                    //        xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                    //        xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                    //        rowExcel++;
                    //    }
                    //}
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {
                        xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                        xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                        xlWorkSheet.Cells[rowExcel, 3] = "RENOVACIÓN";
                        string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                        if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                        else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                        xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                        xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                        xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                        xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                        rowExcel++;
                    }
                }
                //Peticion 3 = Cambio de clase
                if (peticion == 3)
                {
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    string licencia2 = (rangeConduce.Cells[i, 3] as Excel.Range).Value2.ToString();
                    string[] resultado2 = licencia2.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {
                        bool flag = true;
                        for (int k = 0; k < resultado2.Length; k++)
                        {
                            if (resultado[j].Trim() == resultado2[k].Trim()) flag = false;
                        }
                        if (flag == true)
                        {
                            xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                            xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                            xlWorkSheet.Cells[rowExcel, 3] = "CAMBIO DE CLASE";
                            string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                            if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                            else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                            xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                            xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                            xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                            xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                            rowExcel++;
                        }
                    }
                }
                //Peticion 4 = Extension Clase
                if (peticion == 4 || peticion == 7)
                {
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    string licencia2 = (rangeConduce.Cells[i, 3] as Excel.Range).Value2.ToString();
                    string[] resultado2 = licencia2.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {
                        bool flag = true;
                        for (int k = 0; k < resultado2.Length; k++)
                        {
                            if (resultado[j].Trim() == resultado2[k].Trim()) flag = false;
                        }
                        if (flag == true)
                        {
                            xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                            xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                            xlWorkSheet.Cells[rowExcel, 3] = "EXTENSIÓN DE CLASE";
                            string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                            if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                            else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                            xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                            xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                            xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                            xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                            rowExcel++;
                        }
                    }
                }
                //Peticion 5 = Duplicado
                if (peticion == 5)
                {
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {
                        xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                        xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                        xlWorkSheet.Cells[rowExcel, 3] = "DUPLICADO";
                        string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                        if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                        else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                        xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                        xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                        xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                        xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                        rowExcel++;
                    }
                }

                //Peticion 6 = Cambio de domicilio
                if (peticion == 6)
                {
                    string licencia = (rangeConduce.Cells[i, 2] as Excel.Range).Value2.ToString();
                    string[] stringSeparators = new string[] { "," };
                    string[] resultado = licencia.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < resultado.Length; j++)
                    {

                        xlWorkSheet.Cells[rowExcel, 1] = DateTime.FromOADate(((rangeConduce.Cells[i, 1] as Excel.Range).Value2)).Month.ToString();
                        xlWorkSheet.Cells[rowExcel, 2] = ClaseLicencia(Convert.ToInt32(resultado[j]));
                        xlWorkSheet.Cells[rowExcel, 3] = "CAMBIO DE DOMICILIO";
                        string genero = (rangeConduce.Cells[i, 5] as Excel.Range).Value2.ToString();
                        if (genero.Equals("M")) xlWorkSheet.Cells[rowExcel, 4] = "MASCULINO";
                        else xlWorkSheet.Cells[rowExcel, 4] = "FEMENINO";
                        xlWorkSheet.Cells[rowExcel, 5].NumberFormat = "@";
                        xlWorkSheet.Cells[rowExcel, 5] = SetFecha(DateTime.FromOADate(((rangeConduce.Cells[i, 6] as Excel.Range).Value2)));
                        xlWorkSheet.Cells[rowExcel, 6] = string.Format("{0} {1}", (rangeConduce.Cells[i, 7] as Excel.Range).Value2.ToString(), (rangeConduce.Cells[i, 8] as Excel.Range).Value2.ToString());
                        xlWorkSheet.Cells[rowExcel, 7] = (rangeConduce.Cells[i, 9] as Excel.Range).Value2.ToString();
                        rowExcel++;
                    }
                }

            }

            xlWorkBook.SaveAs("D:\\INE\\Licencias\\Ine\\2020\\Final-Licencias.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            Marshal.ReleaseComObject(xlWorkBookConduce);
            Marshal.ReleaseComObject(xlWorkBookConduce);
            Marshal.ReleaseComObject(xlAppConduce);
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


        }

        public string ClaseLicencia(int clase)
        {
            switch (clase)
            {
                case 0:
                    return "B";

                case 1:
                    return "A1";

                case 2:
                    return "A2";

                case 3:
                    return "A1";

                case 4:
                    return "A2";

                case 5:
                    return "A3";

                case 6:
                    return "A4";

                case 7:
                    return "A5";

                case 8:
                    return "C";

                case 9:
                    return "D";

                case 10:
                    return "E";

                case 11:
                    return "F";

                case 12:
                    return "A2";
            }
            return "";
        }

        public string SetFecha(DateTime fecha)
        {
            return fecha.ToString("dd-MM-yyyy");
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\Estadistica\\Estadistica2.txt");
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;
            int rCnt;
            int rw = 0;
            int cl = 0;
            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Estadistica\\Libro1.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            for (int i = 1; i <= 1; i++)
            {
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);
                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;
                int normal = 0;
                int condicionado = 0;
                //Llenar filas
                for (rCnt = 2; rCnt <= 3860; rCnt++)
                {
                    DateTime fechaInicial = DateTime.FromOADate(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    DateTime fechaFinal = DateTime.FromOADate(((range.Cells[rCnt, 3] as Excel.Range).Value2));
                    if (fechaFinal.Year - fechaInicial.Year >= 6) normal++;
                    else condicionado++;
                }
                file.WriteLine("Licencias no profesionales clase B y/o C otorgadas en la Municipalidad de Puchuncaví");
                file.WriteLine("Licencias con vigencia por 6 años = " + normal.ToString());
                file.WriteLine("Licencias con vigencias por menos de 6 años = " + condicionado.ToString());
                Marshal.ReleaseComObject(xlWorkSheet2);
            }
            file.Close();
            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();
            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);
        }

        private void BtnDiputado2019_Click(object sender, RoutedEventArgs e)
        {
            //Crear archivo para escribir estadistica
            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\lICENCIAS\\Estadistica\\Estadistica2019.txt");
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }
            Excel.Application xlApp2;
            Excel.Workbook xlWorkBook2;
            Excel.Worksheet xlWorkSheet2;
            Excel.Range range;
            int rCnt;
            int rw = 0;
            int cl = 0;
            //Listado con edades por año desde al 2014 al 2018 (Solicitado por diputado)
            List<AuxLicencia> edad2014 = new List<AuxLicencia>() { new AuxLicencia() { Edad = 18, Cantidad = 0 } };
            List<AuxLicencia> edad2015 = new List<AuxLicencia>() { new AuxLicencia() { Edad = 18, Cantidad = 0 } };
            List<AuxLicencia> edad2016 = new List<AuxLicencia>() { new AuxLicencia() { Edad = 18, Cantidad = 0 } };
            List<AuxLicencia> edad2017 = new List<AuxLicencia>() { new AuxLicencia() { Edad = 18, Cantidad = 0 } };
            List<AuxLicencia> edad2018 = new List<AuxLicencia>() { new AuxLicencia() { Edad = 18, Cantidad = 0 } };
            //Abrir archivo Excel con datos
            xlApp2 = new Excel.Application();
            xlWorkBook2 = xlApp2.Workbooks.Open("D:\\lICENCIAS\\Estadistica\\DatosLicencias.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            int maximo = Convert.ToInt32(TxtMax.Text);
            for (int i = 1; i <= 1; i++)
            {
                //Seleccionar hoja de trabajo en Excel
                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(i);
                range = xlWorkSheet2.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;                
                for (rCnt = 2; rCnt <= maximo; rCnt++)
                {
                    //Obtener datos por nombre
                    DateTime fechaObtencion = DateTime.FromOADate(((range.Cells[rCnt, 1] as Excel.Range).Value2));
                    string licencia = ((range.Cells[rCnt, 2] as Excel.Range).Value2).ToString();
                    DateTime fechaNacimiento = DateTime.FromOADate(((range.Cells[rCnt, 3] as Excel.Range).Value2));
                    //Seleccionar solo las clase B
                    if (licencia.IndexOf("B") != -1)
                    {
                        // Año 2014
                        if (fechaObtencion.Year == 2014) {
                            int edad = Convert.ToInt32(((fechaObtencion - fechaNacimiento).Days / 365));
                            bool flag = false;
                            foreach (AuxLicencia item in edad2014)
                            {
                                if (item.Edad == edad) {
                                    flag = true;
                                    item.Cantidad = item.Cantidad + 1;
                                }
                            }
                            if (flag == false) {
                                AuxLicencia auxiliar = new AuxLicencia();
                                auxiliar.Edad = edad;
                                auxiliar.Cantidad = 1;
                                edad2014.Add(auxiliar);
                            }
                        }
                        // Año 2015
                        if (fechaObtencion.Year == 2015)
                        {
                            int edad = Convert.ToInt32(((fechaObtencion - fechaNacimiento).Days / 365));
                            bool flag = false;
                            foreach (AuxLicencia item in edad2015)
                            {
                                if (item.Edad == edad)
                                {
                                    flag = true;
                                    item.Cantidad = item.Cantidad + 1;
                                }
                            }
                            if (flag == false)
                            {
                                AuxLicencia auxiliar = new AuxLicencia();
                                auxiliar.Edad = edad;
                                auxiliar.Cantidad = 1;
                                edad2015.Add(auxiliar);
                            }
                        }
                        // Año 2016
                        if (fechaObtencion.Year == 2016)
                        {
                            int edad = Convert.ToInt32(((fechaObtencion - fechaNacimiento).Days / 365));
                            bool flag = false;
                            foreach (AuxLicencia item in edad2016)
                            {
                                if (item.Edad == edad)
                                {
                                    flag = true;
                                    item.Cantidad = item.Cantidad + 1;
                                }
                            }
                            if (flag == false)
                            {
                                AuxLicencia auxiliar = new AuxLicencia();
                                auxiliar.Edad = edad;
                                auxiliar.Cantidad = 1;
                                edad2016.Add(auxiliar);
                            }
                        }
                        // Año 2017
                        if (fechaObtencion.Year == 2017)
                        {
                            int edad = Convert.ToInt32(((fechaObtencion - fechaNacimiento).Days / 365));
                            bool flag = false;
                            foreach (AuxLicencia item in edad2017)
                            {
                                if (item.Edad == edad)
                                {
                                    flag = true;
                                    item.Cantidad = item.Cantidad + 1;
                                }
                            }
                            if (flag == false)
                            {
                                AuxLicencia auxiliar = new AuxLicencia();
                                auxiliar.Edad = edad;
                                auxiliar.Cantidad = 1;
                                edad2017.Add(auxiliar);
                            }
                        }
                        // Año 2018
                        if (fechaObtencion.Year == 2018)
                        {
                            int edad = Convert.ToInt32(((fechaObtencion - fechaNacimiento).Days / 365));
                            bool flag = false;
                            foreach (AuxLicencia item in edad2018)
                            {
                                if (item.Edad == edad)
                                {
                                    flag = true;
                                    item.Cantidad = item.Cantidad + 1;
                                }
                            }
                            if (flag == false)
                            {
                                AuxLicencia auxiliar = new AuxLicencia();
                                auxiliar.Edad = edad;
                                auxiliar.Cantidad = 1;
                                edad2018.Add(auxiliar);
                            }
                        }
                    }
                }
                // Agregar al archivo plano
                file.WriteLine(string.Format("La cantidad de licencias clase B por año es "));
                file.WriteLine("");
                file.WriteLine("Año 2014");
                int suma = 0;
                foreach (AuxLicencia item in edad2014)
                {
                    file.WriteLine(string.Format("Edad " + item.Edad + ": " + item.Cantidad));
                    suma = suma + item.Cantidad;
                }
                file.WriteLine(string.Format("Suma del año: "+suma));
                file.WriteLine("");
                file.WriteLine("Año 2015");
                suma = 0;
                foreach (AuxLicencia item in edad2015)
                {
                    file.WriteLine(string.Format("Edad " + item.Edad + ": " + item.Cantidad)); suma = suma + item.Cantidad;
                }
                file.WriteLine(string.Format("Suma del año: " + suma));
                file.WriteLine("");
                file.WriteLine("Año 2016");
                suma = 0;
                foreach (AuxLicencia item in edad2016)
                {
                    file.WriteLine(string.Format("Edad " + item.Edad + ": " + item.Cantidad));
                    suma = suma + item.Cantidad;
                }
                file.WriteLine(string.Format("Suma del año: " + suma));
                file.WriteLine("");
                file.WriteLine("Año 2017");
                suma = 0;
                foreach (AuxLicencia item in edad2017)
                {
                    file.WriteLine(string.Format("Edad " + item.Edad + ": " + item.Cantidad));
                    suma = suma + item.Cantidad;
                }
                file.WriteLine(string.Format("Suma del año: " + suma));
                file.WriteLine("");
                file.WriteLine("Año 2018");
                suma = 0;
                foreach (AuxLicencia item in edad2018)
                {
                    file.WriteLine(string.Format("Edad " + item.Edad + ": " + item.Cantidad));
                    suma = suma + item.Cantidad;
                }
                file.WriteLine(string.Format("Suma del año: " + suma));
            }
            file.Close();
            xlWorkBook2.Close(true, null, null);
            xlApp2.Quit();
            Marshal.ReleaseComObject(xlWorkBook2);
            Marshal.ReleaseComObject(xlApp2);
        }
    }
}
