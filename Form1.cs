using System;
using System.Data;
using System.Data.OleDb;// para conectarse a bases de datos
using System.Windows.Forms;

namespace Toda_Las_posibilidades
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string data0 = "";
        string data1 = "";

        int Num0 = 0;
        int Num1 = 0;
        int Num2 = 0;
        int Num3 = 0;
        int Num4 = 0;
        int Num5 = 0;

        double Num6 = 0;
        double Sumador = 0;

        int X1, X2, X3, X4 = 0;
        bool vs = true;
        bool cs = true;
        private void button2_Click(object sender, EventArgs e)
        {
            vs = true;
            X3 = 0;
            while (vs)
            {
                data0 = dataGridView1.Rows[X1].Cells[0].Value.ToString();

                data1 = dataGridView1.Rows[X1].Cells[1].Value.ToString();
                Num1 = int.Parse(data1);

                while (vs)
                {
                    data0 = dataGridView1.Rows[X2].Cells[0].Value.ToString();
                    Num2 = int.Parse(data0);
                    data1 = dataGridView1.Rows[X2].Cells[1].Value.ToString();
                    Num3 = int.Parse(data1);

                    if ((Num0 == Num2) && (Num1 == Num3))
                    {
                        X3++;
                    }

                    if ((Num1 == Num2) && (Num0 == Num3))
                    {
                        X3++;
                    }

                    X2++;
                    if (data0 == "100") { vs = false; }
                }
                vs = true;
                X1++;
                X2 = 0;
                if (data0 == "100") { vs = false; }
                dataGridView2.Rows[0].Cells[2].Value = X3;
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            while (X1 <= 9999)
            {
                data0 = dataGridView2.Rows[X1].Cells[0].Value.ToString();
                Num0 = int.Parse(data0);
                data1 = dataGridView2.Rows[X1].Cells[1].Value.ToString();
                Num1 = int.Parse(data1);
                while (vs)
                {
                    data0 = dataGridView1.Rows[X2].Cells[1].Value.ToString();
                    Num2 = int.Parse(data0);
                    data1 = dataGridView1.Rows[X2].Cells[2].Value.ToString();
                    Num3 = int.Parse(data1);
                    if ((Num0 == Num2) && (Num1 == Num3))
                    {
                        dataGridView2.Rows[X1].Cells[2].Value = "X";
                        vs = false;
                    }
                    else { dataGridView2.Rows[X1].Cells[2].Value = "Y"; }

                    if ((Num1 == Num2) && (Num0 == Num3))
                    {
                        dataGridView2.Rows[X1].Cells[2].Value = "X";
                        vs = false;
                    }
                    else { dataGridView2.Rows[X1].Cells[2].Value = "Y"; }

                    X2++;
                    if (data0 == "100") { vs = false; }
                }
                vs = true;
                X1++;
                X2 = 0;
            }

            //X1 = 0;
            //while (X1 < 9999)
            //{
            //    data0 = dataGridView2.Rows[X1].Cells[2].Value.ToString();
            //    if (data0 != "x")
            //    {
            //        dataGridView3.Rows.Add();
            //        data0 = dataGridView2.Rows[X2].Cells[0].Value.ToString();
            //        data1 = dataGridView2.Rows[X2].Cells[1].Value.ToString();
            //        dataGridView3.Rows[X1].Cells[0].Value = data0;
            //        dataGridView3.Rows[X1].Cells[1].Value = data1;
            //        X2++;
            //    }
            //    X1++;
            //}
            //X1 = 0;
        }
        DataView ImportarDatos(string nombrearchivo) //COMO PARAMETROS OBTENEMOS EL NOMBRE DEL ARCHIVO A IMPORTAR
        {

            //UTILIZAMOS 12.0 DEPENDIENDO DE LA VERSION DEL EXCEL, EN CASO DE QUE LA VERSIÓN QUE TIENES ES INFERIOR AL DEL 2013, CAMBIAR A EXCEL 8.0 Y EN VEZ DE
            //ACE.OLEDB.12.0 UTILIZAR LO SIGUIENTE (Jet.Oledb.4.0)
            string conexion = string.Format("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = {0}; Extended Properties = 'Excel 12.0;'", nombrearchivo);
            nombrearchivo = "Regristro Actualizado";
            OleDbConnection conector = new OleDbConnection(conexion);

            conector.Open();

            //DEPENDIENDO DEL NOMBRE QUE TIENE LA PESTAÑA EN TU ARCHIVO EXCEL COLOCAR DENTRO DE LOS []
            OleDbCommand consulta = new OleDbCommand("select * from [Hoja1$]", conector);

            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta
            };

            DataSet ds = new DataSet();

            adaptador.Fill(ds);

            conector.Close();

            return ds.Tables[0].DefaultView;


        }
        private void button3_Click(object sender, EventArgs e)
        {
            Sumador1();
            dataGridView3.Rows[0].Cells[0].Value = Sumador;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            X1 = 0;
            cs = true;
            while (cs)
            {
                data0 = dataGridView1.Rows[X1].Cells[3].Value.ToString();
                if (data0 == "100") { cs = false; break; }
                if (data0 == "X") { X2++; } else { X3++; }
                X1++;

            }
            dataGridView3.Rows[1].Cells[0].Value = X2 + "X";
            dataGridView3.Rows[2].Cells[0].Value = X3 + "Y";
            X1 = 0;
            X2 = 0;
            X3 = 0;
            cs = true;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //DE ESTA MANERA FILTRAMOS TODOS LOS ARCHIVOS EXCEL EN EL NAVEGADOR DE ARCHIVOS
                Filter = "Excel | *.xls;*.xlsx;",

                //AQUÍ INDICAMOS QUE NOMBRE TENDRÁ EL NAVEGADOR DE ARCHIVOS COMO TITULO
                Title = "Seleccionar Archivo"
            };
            //EN CASO DE SELECCIONAR EL ARCHIVO, ENTONCES PROCEDEMOS A ABRIR EL ARCHIVO CORRESPONDIENTE
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dataGridView1.DataSource = ImportarDatos(openFileDialog.FileName);
            }
            dataGridView1.Columns.Add("Column10", "X4");
            dataGridView1.Columns.Add("Column11", "Numeros");
            dataGridView2.Rows.Add(9999);
            dataGridView3.Rows.Add(4);
            for (int i = 0, k = 0; i < 100; i++)
            {
                for (int f = 0; f < 100; f++, k++)
                {
                    dataGridView2.Rows[k].Cells[1].Value = f;
                    dataGridView2.Rows[k].Cells[0].Value = i;
                }
            }
            X1 = 0;
            while (true)
            {
                data0 = dataGridView1.Rows[X1].Cells[0].Value.ToString();
                dataGridView1.Rows[X1].Cells[4].Value = X1;
                X1++;
                if (data0 == "100") 
                {
                    dataGridView1.Rows[X1].Cells[4].Value = X1;
                    cs = false; break;          
                }
            }
            X1 = 0;
        }
        void Sumador1(int colum = 0)
        {
            data0 = "0";
            while (cs)
            {
                data0 = dataGridView1.Rows[X1].Cells[0].Value.ToString();
                Num0 = int.Parse(data0);
                if (Num0 == 100)
                {
                    X1 = 0;
                    Validacion();
                    if (X4 > 0)
                    {
                        Sumador1();
                    }
                    return;
                }
                Num1 = Num0 - 30;// Limite inferior de entrada
                if (Num1 < 0) { Num1 = 0; }
                Num2 = Num0 + 30;// Limite superior de entrada
                if (Num2 > 99) { Num2 = 99; }
                X1++;

                data0 = dataGridView1.Rows[X1].Cells[0].Value.ToString();
                Num3 = int.Parse(data0);

                if ((Num3 >= Num1) && (Num3 <= Num2))// Si esta dentro del rango establecido
                {
                    dataGridView1.Rows[X1 - 1].Cells[3].Value = "X";
                    while (cs)
                    {
                        Num4 = Num3 - 5;// Limite inferior de salida
                        Num5 = Num3 + 5;// Limite superior de salida
                        Num6 = Sumador + Num0;
                            if (Num6 <= Num4) { Sumador += 0.5; }
                            if (Num6 >= Num5) { Sumador -= 0.5; }
                            if ((Num6 >= Num4) && (Num6 <= Num5)) { cs = false; } else { cs = true; }// Si cumple con los limites, sale del bucle
                        dataGridView3.Rows[0].Cells[1].Value = X4;
                    }
                }
                else { dataGridView1.Rows[X1 - 1].Cells[3].Value = "Y"; }
                cs = true;
                dataGridView1.Rows[X1 + 1].Cells[3].Value = "100";
            }
        }
        void Validacion()
        {
            X2 = 0;
            cs = true;
            while (cs)
            {
                data0 = dataGridView1.Rows[X2].Cells[0].Value.ToString();
                Num0 = int.Parse(data0);
                if (Num0 == 100)
                {
                    cs = false;
                    break;
                }
                Num1 = Num0 - 30;// Limite inferior de entrada
                Num2 = Num0 + 30;// Limite superior de entrada
                X2++;
                data0 = dataGridView1.Rows[X2].Cells[0].Value.ToString();
                Num3 = int.Parse(data0);
                if ((Num3 >= Num1) && (Num3 <= Num2))// Si esta dentro del rango establecido
                {
                    Num4 = Num3 - 5;// Limite inferior de salida
                    Num5 = Num3 + 5;// Limite superior de salida
                    Num6 = Sumador + Num3;
                    if ((Num6 >= Num4) && (Num6 <= Num5)) { } else { X4++; }
                }
            }
        }
    }
}
