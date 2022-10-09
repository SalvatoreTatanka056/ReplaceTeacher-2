using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using System.Data;
using System.Globalization;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Drawing;
using System.Reflection.PortableExecutable;
using System.Numerics;
using System.Configuration;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using System.Security.Policy;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ReplaceTeacher
{
    public struct structureAssentiDisponibili
    {
        public structureAssentiDisponibili
        (
        string nome_docente,
        string prima,
        string seconda,
        string terza,
        string quarta,
        string quinta,
        string sesta,
        string settima,
        string ottava,
        bool visto)
        {
            Nome_docente = nome_docente;
            Prima = prima;
            Seconda = seconda;
            Terza = terza;
            Quarta = quarta;
            Quinta = quinta;
            Sesta = sesta;
            Settima = settima;
            Ottava = ottava;
            Visto1 = false;
            Visto2 = false;
            Visto3 = false;
            Visto4 = false;
            Visto5 = false;
            Visto6 = false;
            Visto7 = false;
            Visto8 = false;
        }

        public string Nome_docente { get; init; }
        public string Prima { get; init; }
        public string Seconda { get; init; }
        public string Terza { get; init; }
        public string Quarta { get; init; }
        public string Quinta { get; init; }
        public string Sesta { get; init; }
        public string Settima { get; init; }
        public string Ottava { get; init; }
        public bool Visto1 { get; set; }
        public bool Visto2 { get; set; }
        public bool Visto3 { get; set; }
        public bool Visto4 { get; set; }
        public bool Visto5 { get; set; }
        public bool Visto6 { get; set; }
        public bool Visto7 { get; set; }
        public bool Visto8 { get; set; }
    }

    public partial class frmMain : Form
    {
        List<structureAssentiDisponibili> listDisponibili;
        List<structureAssentiDisponibili> listAssenti;
        OracleConnection conn;

        public bool m_bdtgridview = false;

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {

            string connString = "DATA SOURCE=localhost:1522/xepdb1;" + "PERSIST SECURITY INFO=True;USER ID=system; password=system; Pooling = False; ";
            conn = new OracleConnection();
            conn.ConnectionString = connString;
            conn.Open();

            OracleCommand cmd = conn.CreateCommand();
            cmd.CommandText = "select * from docenti";
            OracleDataReader reader = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(reader);

            dataGridView1.DataSource = dt;
            dataGridView1.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView1.Columns[0].Width = 0;
            cmd.Dispose();
            reader.Dispose();

            OracleCommand cmdC = conn.CreateCommand();
            //cmdC.CommandText = @"select * from  calendario_lezioni";
            cmdC.CommandText = @"select * from calendario_lezioni";
            OracleDataReader readerC = cmdC.ExecuteReader();
            DataTable dtC = new DataTable();
            dtC.Load(readerC);

            dataGridView3.DataSource = dtC;
            dataGridView3.EditMode = DataGridViewEditMode.EditOnEnter;
            dataGridView3.Columns[1].Width = 0;
            dataGridView3.Columns[2].Width = 0;
            cmdC.Dispose();
            readerC.Dispose();

        }

        private void ElencoAssenti()
        {
            OracleCommand cmd = conn.CreateCommand();
            cmd.CommandText = "select * from calendario_lezioni where nome_docente  in(select nome_docente from docenti where assente = 'x') and giorno = TO_CHAR(TO_DATE('" + dateTimePicker1.Value.ToString() + "','DD/MM/YYYY hh24:mi:ss'), 'D')";
            string str = cmd.CommandText;
            OracleDataReader reader = cmd.ExecuteReader();

            int custIdCol0 = reader.GetOrdinal("Giorno");
            int custIdCol1 = reader.GetOrdinal("Nome_docente");
            int custIdCol2 = reader.GetOrdinal("prima");
            int custIdCol3 = reader.GetOrdinal("seconda");
            int custIdCol4 = reader.GetOrdinal("terza");
            int custIdCol5 = reader.GetOrdinal("quarta");
            int custIdCol6 = reader.GetOrdinal("quinta");
            int custIdCol7 = reader.GetOrdinal("sesta");
            int custIdCol8 = reader.GetOrdinal("settima");
            int custIdCol9 = reader.GetOrdinal("ottava");

            listAssenti = new List<structureAssentiDisponibili>();

            while (reader.Read())
            {
                string p1 = reader.GetString(custIdCol1);
                string p2 = "", p3 = "", p4 = "", p5 = "", p6 = "", p7 = "", p8 = "", p9 = "";

                if (reader.IsDBNull(custIdCol2) == false)
                {
                    p2 = reader.GetString(custIdCol2);
                }
                if (reader.IsDBNull(custIdCol3) == false)
                {
                    p3 = reader.GetString(custIdCol3);
                }
                if (reader.IsDBNull(custIdCol4) == false)
                {
                    p4 = reader.GetString(custIdCol4); ;
                }
                if (reader.IsDBNull(custIdCol5) == false)
                {
                    p5 = reader.GetString(custIdCol5);
                }
                if (reader.IsDBNull(custIdCol6) == false)
                {
                    p6 = reader.GetString(custIdCol6);
                }
                if (reader.IsDBNull(custIdCol7) == false)
                {
                    p7 = reader.GetString(custIdCol7);
                }
                if (reader.IsDBNull(custIdCol8) == false)
                {
                    p8 = reader.GetString(custIdCol8);
                }
                if (reader.IsDBNull(custIdCol9) == false)
                {
                    p9 = reader.GetString(custIdCol9);
                }

                structureAssentiDisponibili assentiDisponibili = new structureAssentiDisponibili(reader.GetString(custIdCol1),
                                            p2, p3, p4, p5, p6, p7, p8, p9, false);

                listAssenti.Add(assentiDisponibili);
            }

            cmd.Dispose();
            reader.Dispose();
        }

        private void ElencoDisponibili()
        {
            OracleCommand cmd = conn.CreateCommand();
            cmd.CommandText = "select * from calendario_lezioni where nome_docente not in(select nome_docente from docenti where assente = 'x') and giorno = TO_CHAR(TO_DATE('" + dateTimePicker1.Value.ToString() + "','DD/MM/YYYY hh24:mi:ss'), 'D')";

            OracleDataReader reader = cmd.ExecuteReader();

            int custIdCol1 = reader.GetOrdinal("Nome_docente");
            int custIdCol2 = reader.GetOrdinal("prima");
            int custIdCol3 = reader.GetOrdinal("seconda");
            int custIdCol4 = reader.GetOrdinal("terza");
            int custIdCol5 = reader.GetOrdinal("quarta");
            int custIdCol6 = reader.GetOrdinal("quinta");
            int custIdCol7 = reader.GetOrdinal("sesta");
            int custIdCol8 = reader.GetOrdinal("settima");
            int custIdCol9 = reader.GetOrdinal("ottava");

            listDisponibili = new List<structureAssentiDisponibili>();

            while (reader.Read())
            {
                string p1 = reader.GetString(custIdCol1);
                string p2 = "", p3 = "", p4 = "", p5 = "", p6 = "", p7 = "", p8 = "", p9 = "";

                if (reader.IsDBNull(custIdCol2) == false)
                {
                    p2 = reader.GetString(custIdCol2);
                }
                if (reader.IsDBNull(custIdCol3) == false)
                {
                    p3 = reader.GetString(custIdCol3);
                }
                if (reader.IsDBNull(custIdCol4) == false)
                {
                    p4 = reader.GetString(custIdCol4); ;
                }
                if (reader.IsDBNull(custIdCol5) == false)
                {
                    p5 = reader.GetString(custIdCol5);
                }
                if (reader.IsDBNull(custIdCol6) == false)
                {
                    p6 = reader.GetString(custIdCol6);
                }
                if (reader.IsDBNull(custIdCol7) == false)
                {
                    p7 = reader.GetString(custIdCol7);
                }
                if (reader.IsDBNull(custIdCol8) == false)
                {
                    p8 = reader.GetString(custIdCol8);
                }
                if (reader.IsDBNull(custIdCol9) == false)
                {
                    p9 = reader.GetString(custIdCol9);
                }

                structureAssentiDisponibili assentiDisponibili = new structureAssentiDisponibili(reader.GetString(custIdCol1), p2, p3, p4, p5, p6, p7, p8, p9, false);

                listDisponibili.Add(assentiDisponibili);
            }
            cmd.Dispose();
            reader.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ElencoAssenti();
            ElencoDisponibili();


            m_bdtgridview = false;

            DataTable table = new DataTable();
            table.Columns.Add("Giorno", typeof(string));
            table.Columns.Add("Nome_Assente", typeof(string));
            table.Columns.Add("1°", typeof(string));
            table.Columns.Add("Nome_Docente_1", typeof(string));
            table.Columns.Add("2°", typeof(string));
            table.Columns.Add("Nome_Docente_2", typeof(string));
            table.Columns.Add("3°", typeof(string));
            table.Columns.Add("Nome_Docente_3", typeof(string));
            table.Columns.Add("4°", typeof(string));
            table.Columns.Add("Nome_Docente_4", typeof(string));
            table.Columns.Add("5°", typeof(string));
            table.Columns.Add("Nome_Docente_5", typeof(string));
            table.Columns.Add("6°", typeof(string));
            table.Columns.Add("Nome_Docente_6", typeof(string));
            table.Columns.Add("7°", typeof(string));
            table.Columns.Add("Nome_Docente_7", typeof(string));
            table.Columns.Add("8°", typeof(string));
            table.Columns.Add("Nome_Docente_8", typeof(string));

            foreach (structureAssentiDisponibili itemA in listAssenti)
            {
                DataRow myDataRow;
                myDataRow = table.NewRow();
                myDataRow["Giorno"] = dateTimePicker1.Value.DayOfWeek.ToString();
                myDataRow["Nome_Assente"] = itemA.Nome_docente;
                myDataRow["1°"] = itemA.Prima;
                myDataRow["Nome_Docente_1"] = "NS";
                myDataRow["2°"] = itemA.Seconda;
                myDataRow["Nome_docente_2"] = "NS";
                myDataRow["3°"] = itemA.Terza;
                myDataRow["Nome_docente_3"] = "NS";
                myDataRow["4°"] = itemA.Quarta;
                myDataRow["Nome_docente_4"] = "NS";
                myDataRow["5°"] = itemA.Quinta;
                myDataRow["Nome_docente_5"] = "NS";
                myDataRow["6°"] = itemA.Sesta;
                myDataRow["Nome_docente_6"] = "NS";
                myDataRow["7°"] = itemA.Settima;
                myDataRow["Nome_docente_7"] = "NS";
                myDataRow["8°"] = itemA.Ottava;
                myDataRow["Nome_docente_8"] = "NS";

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto1 == false)
                    {
                        if (itemA.Prima.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Prima.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Prima.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_1"] = itemD.Nome_docente;

                                    itemD.Visto1 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto2 == false)
                    {
                        if (itemA.Seconda.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Seconda.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Seconda.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_2"] = itemD.Nome_docente;

                                    itemD.Visto2 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto3 == false)
                    {

                        if (itemA.Terza.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Terza.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Terza.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_3"] = itemD.Nome_docente;

                                    itemD.Visto3 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto4 == false)
                    {

                        if (itemA.Quarta.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Quarta.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Quarta.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_4"] = itemD.Nome_docente;

                                    itemD.Visto4 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto5 == false)
                    {

                        if (itemA.Quinta.Length != 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Quinta.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Quinta.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_5"] = itemD.Nome_docente;

                                    itemD.Visto5 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {
                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto6 == false)
                    {
                        if (itemA.Sesta.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Sesta.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Sesta.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_6"] = itemD.Nome_docente;

                                    itemD.Visto6 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {

                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto7 == false)
                    {

                        if (itemA.Settima.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Settima.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Settima.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_7"] = itemD.Nome_docente;

                                    itemD.Visto7 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < listDisponibili.Count(); i++)
                {

                    structureAssentiDisponibili itemD = listDisponibili[i];

                    if (itemD.Visto8 == false)
                    {
                        if (itemA.Ottava.Length == 0)
                        {
                            break;
                        }
                        else
                        {
                            if (itemD.Ottava.Length == 0)
                                continue;
                            else
                            {
                                if (itemD.Ottava.Equals("P"))
                                {
                                    myDataRow["Nome_Docente_8"] = itemD.Nome_docente;

                                    itemD.Visto8 = true;
                                    listDisponibili[i] = itemD;
                                }
                            }

                        }
                    }
                }

                table.Rows.Add(myDataRow);
            }
            dataGridView2.DataSource = table;

            /* for (int x = 0; x < dataGridView2.Rows.Count; x++)
             {
                 for (int y = 0; y < dataGridView2.Rows[x].Cells.Count; y++)
                 {
                     if (dataGridView2.Rows[x].Cells[y].Value != null)
                     {
                         if (dataGridView2.Rows[x].Cells[y].Value.Equals("NS"))
                         {
                             int col = dataGridView2.Rows[x].Cells[y].ColumnIndex;
                             dataGridView2.Columns[col].Width = 0;
                         }
                     }
                 }
             }*/

            m_bdtgridview = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var dataTable = ((DataTable)dataGridView1.DataSource).GetChanges();
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    switch (row.RowState)
                    {
                        case DataRowState.Added:
                            // to do INSERT QUERY
                            break;
                        case DataRowState.Deleted:
                            // to do DELETE QUERY
                            break;
                        case DataRowState.Modified:
                            OracleCommand cmd = conn.CreateCommand();
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = "UPDATE Docenti SET Assente = :param1 where NOME_DOCENTE = :param2";
                            cmd.Parameters.Add("param1", row["Assente"]);
                            cmd.Parameters.Add("param2", row["Nome_docente"]);
                            cmd.ExecuteNonQuery();
                            cmd.Dispose();
                            break;
                    }
                }

              ((DataTable)dataGridView1.DataSource).AcceptChanges();
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ExportGridToword();
            /*printPreviewDialog1.ShowDialog();
            if (printDialog1.ShowDialog() == DialogResult.OK)
            {

                //printDocument1.Print();
            }*/
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Pink;

            try
            {
                if (e.Value != null)
                {
                    if (e.Value.Equals("P"))
                        e.CellStyle.BackColor = Color.LightGreen;
                    else if (!e.Value.Equals("P") && !e.Value.Equals("NS"))
                        e.CellStyle.BackColor = Color.LightBlue;

                    if (e.ColumnIndex == 1)
                        e.CellStyle.BackColor = Color.AntiqueWhite;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Errore: " + err.Message);
            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

            var dataTable = ((DataTable)dataGridView3.DataSource).GetChanges();
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                foreach (DataRow row in dataTable.Rows)
                {
                    switch (row.RowState)
                    {
                        case DataRowState.Added:
                            // to do INSERT QUERY
                            break;
                        case DataRowState.Deleted:
                            // to do DELETE QUERY
                            break;
                        case DataRowState.Modified:
                            OracleCommand cmd = conn.CreateCommand();
                            cmd.CommandType = CommandType.Text;

                            cmd.CommandText = @"UPDATE CALENDARIO_LEZIONI " +
                                             "SET GIORNO = :param1, " +
                                             "PRIMA = :param2, " +
                                             "SECONDA = :param3, " +
                                             "TERZA = :param4, " +
                                             "QUARTA = :param5, " +
                                             "QUINTA = :param6, " +
                                             "SESTA = :param7, " +
                                             "SETTIMA = :param8, " +
                                             "OTTAVA = :param9 , " +
                                             "NOME_DOCENTE = :param10 " +
                                             "WHERE NOME_DOCENTE = :param10 and giorno = :param1";

                            cmd.Parameters.Add("param1", row["GIORNO"]);
                            cmd.Parameters.Add("param2", row["PRIMA"]);
                            cmd.Parameters.Add("param3", row["SECONDA"]);
                            cmd.Parameters.Add("param4", row["TERZA"]);
                            cmd.Parameters.Add("param5", row["QUARTA"]);
                            cmd.Parameters.Add("param6", row["QUINTA"]);
                            cmd.Parameters.Add("param7", row["SESTA"]);
                            cmd.Parameters.Add("param8", row["SETTIMA"]);
                            cmd.Parameters.Add("param19", row["OTTAVA"]);
                            cmd.Parameters.Add("param10", row["NOME_DOCENTE"]);
                            cmd.ExecuteNonQuery();

                            cmd.Dispose();
                            break;
                    }
                }

              ((DataTable)dataGridView3.DataSource).AcceptChanges();
            }

        }

        private void ExportGridToword()
        {

            //Table start.
            string html = "<table cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-size: 9pt;font-family:arial'>";

            //Adding HeaderRow.
            html += "<tr>";
            foreach (DataGridViewColumn column in dataGridView2.Columns)
            {
                html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>" + column.HeaderText + "</th>";
            }
            html += "</tr>";

            //Adding DataRow.
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                html += "<tr>";
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                        html += "<td style='width:120px;border: 1px solid #ccc'>" + cell.Value.ToString() + "</td>";
                }
                html += "</tr>";
            }

            //Table end.
            html += "</table>";

            File.WriteAllText(@"DataGridView.htm", html);



            MessageBox.Show("Export Terminato");

            try
            {
                string filename = Path.GetFullPath("DataGridView.htm");
                var uri = new Uri(Path.Combine(Application.StartupPath, "web", filename));

                Process.Start("explorer.exe", filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void splitContainer4_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void splitContainer3_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }
    }
}