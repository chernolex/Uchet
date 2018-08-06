using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Timers;

namespace Uchet
{
    public partial class Form1 : Form
    {
        OleDbConnection conect = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Uchet.accdb;Jet OLEDB:Database Password=777");
        OleDbCommand command = new OleDbCommand();
        string passwordAdmin;
        bool passwordEntered = false;
        List<TextBoxWithColumnIndex> resultTextBoxes = new List<TextBoxWithColumnIndex>();
        System.Timers.Timer time = new System.Timers.Timer(300000);

        public Form1()
        {
            time.Elapsed += OnTimedEvent;
            time.AutoReset = true;
            passwordAdmin = "Perspectiva";
            InitializeComponent();
            conect.Open();
            listView1.Columns[0].Width = listView1.Width - 4;
            listView1.SelectedIndices.Add(0);
            dataGridView1.DefaultCellStyle.DataSourceNullValue = "";
            this.Width += 1;
        }

        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            //textBox1.Text = "";
            Invoke(new MethodInvoker(() => { textBox1.Text = ""; }));
        }

        void CommandExecution(string commandText)
        {
            command = conect.CreateCommand();

            command.CommandText = commandText;
            command.CommandType = CommandType.Text;
            try
            {
                command.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
                command.CommandText = command.CommandText.Replace("\"\"", "0");
                command.ExecuteNonQuery();
            }

            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            DataTable table = new DataTable("Приход");

            //adapter.Fill(table);
            //dataGridView1.DataSource = table.DefaultView;
            //adapter.Update(table);
            RefreshTable();
        }

        void RefreshTable()
        {
            command = conect.CreateCommand();
            string s = "";
            if (listView1.SelectedItems[0].Text == "Итоги")
            {
                s = ReadQuery("queries\\result.query");
            }
            else
            {
                s = "SELECT * FROM " + listView1.SelectedItems[0].Tag;
            }
            command.CommandText = s;
            command.CommandType = CommandType.Text;

            command.ExecuteNonQuery();
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            DataTable table = new DataTable("Приход");

            adapter.Fill(table);
            try
            {
                //dataGridView1.Columns.Clear();
            }
            catch
            {

            }
            dataGridView1.DataSource = table.DefaultView;
            adapter.Update(table);

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].HeaderText = dataGridView1.Columns[i].HeaderText.Replace("\"", "");
            }
            if (listView1.SelectedItems[0].Text == "Итоги")
            {
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[0].Frozen = true;
            }
            else
            {
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                dataGridView1.Columns[0].Visible = false;
            }
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns[0].DefaultCellStyle.BackColor = Color.LightGray;
            try
            {
                dataGridView1.Columns["Итого"].ReadOnly = true;
                dataGridView1.Columns["Итого"].DefaultCellStyle.BackColor = Color.LightGray;
            }
            catch (System.Exception ex)
            {

            }

            foreach (var ri in resultTextBoxes)
            {
                ri.tb.Dispose();
            }
            resultTextBoxes.Clear();

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                try
                {
                    if ((dataGridView1.Columns[i].HeaderText == "Сумма") || (dataGridView1.Columns[i].HeaderText == "Итого") || ((dataGridView1.Columns[i].HeaderText.Length > 6) && (dataGridView1.Columns[i].HeaderText.Remove(7) == "Выручка")) || (dataGridView1.Columns[i].HeaderText.Contains("Приход")))
                    {
                        AddTextBox(i);
                    }
                }
                catch (System.Exception ex)
                {

                }
            }

            //Последняя строка- недоступна для редактирования кроме даты
            if (listView1.SelectedItems[0].Text == "Итоги")
            {
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].ReadOnly = true;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightGray;
                dataGridView1[0, dataGridView1.Rows.Count - 1].ReadOnly = false;
                dataGridView1[0, dataGridView1.Rows.Count - 1].Style.BackColor = Color.White;
            }

            if (listView1.SelectedItems[0].Text == "Итоги")
            {
                for (int i = 2; i < dataGridView1.Columns.Count - 1; i++)
                {

                    if (((i > 3) && (i < 16)) || (i == 18))
                    {
                        dataGridView1.Columns[i].ReadOnly = true;
                        dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGray;
                    }
                    AddTextBox(i);
                }
            }

            int chInd = FindColumn("Проверил");

            if (chInd >= 0)
            {
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    if (dataGridView1[chInd, i].Value.ToString() != "")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.GreenYellow;
                    }
                }
            }

            //dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
            try
            {
                int colInd = 1;
                if (listView1.SelectedItems[0].Text == "Итоги")
                    colInd = 0;
                dataGridView1.CurrentCell = dataGridView1[colInd, dataGridView1.RowCount - 2];
            }
            catch (System.Exception ex)
            {
            	
            }
            
        }

        void AddTextBox(int i)
        {
            TextBox tb = new TextBox();
            tb.ReadOnly = true;
            tb.Left = GetColumnCoord(i);
            tb.Width = dataGridView1.Columns[i].Width;
            tb.Text = GetColumnSum(i).ToString();
            tb.Top = 20;
            this.Controls.Add(tb);
            resultTextBoxes.Add(new TextBoxWithColumnIndex(tb, i));
            dataGridView1.Columns[i].ToolTipText = (resultTextBoxes.Count - 1).ToString();
        }

        double GetColumnSum(int columnIndex)
        {
            double summ = 0;
            int colInd = FindColumn("Проверил");
            int dateCol = 1;
            if(listView1.SelectedItems[0].Text == "Итоги")
                dateCol = 0;

            for (int i = dataGridView1.RowCount - 2; i >=0; i--)
            {
                if (dateTimePicker2.Value.Date.CompareTo(DateTime.Parse(dataGridView1[dateCol, i].Value.ToString())) > -1)
                {
                    if (dateTimePicker1.Value.Date.CompareTo(DateTime.Parse(dataGridView1[dateCol, i].Value.ToString())) > 0)
                        break;
                    if (colInd == -1)
                    {
                        summ += double.Parse(dataGridView1[columnIndex, i].Value.ToString());
                    }
                    else
                        if ((dataGridView1[colInd, i].Value.ToString() != ""))
                            summ += double.Parse(dataGridView1[columnIndex, i].Value.ToString());
                }
            }

            return summ;
        }

        double GetColumnSum(int columnIndex, DataGridView dgv)
        {
            double summ = 0;
            int colInd = FindColumn("Проверил", dgv);

            for (int i = 0; i < dgv.RowCount - 1; i++)
            {
                if (colInd == -1)
                {
                    summ += double.Parse(dgv[columnIndex, i].Value.ToString());
                }
                else
                    if ((dgv[colInd, i].Value.ToString() != ""))
                        summ += double.Parse(dgv[columnIndex, i].Value.ToString());
            }

            return summ;
        }

        int FindColumn(string columnToFind)
        {
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (dataGridView1.Columns[i].HeaderText == columnToFind)
                    return i;
            }

            return -1;
        }

        int FindColumn(string columnToFind, DataGridView dgv)
        {
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                if (dgv.Columns[i].HeaderText == columnToFind)
                    return i;
            }

            return -1;
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            bool isDateFound = false;

            if ((e.ColumnIndex >= 0) && (e.RowIndex >= 0))
            {
                string editableText = dataGridView1[e.ColumnIndex, e.RowIndex].Value.ToString();
                var t = dataGridView1[e.ColumnIndex, e.RowIndex];

                if (CheckRights())
                {
                    int idCol = 0;
                    if (listView1.SelectedItems[0].Text == "Итоги")
                        idCol = 1;
                    if ((dataGridView1[idCol, e.RowIndex].Value.ToString() == "") || (dataGridView1[idCol, e.RowIndex].Value.ToString() == "0"))
                    {
                        if (listView1.SelectedItems[0].Text != "Итоги")
                            CommandExecution("INSERT INTO " + listView1.SelectedItems[0].Tag + " ([Дата]) VALUES (\"" + DateTime.Today.ToShortDateString() + "\")");
                        else
                        {
                            try
                            {
                                DateTime insertingDate = DateTime.Parse(dataGridView1[0, e.RowIndex].Value.ToString());

                                int curRow = 3;
                                DateTime curDate = DateTime.Parse(dataGridView1[0, dataGridView1.RowCount - 3].Value.ToString());
                                while (insertingDate <= curDate)
                                {
                                    if (insertingDate.CompareTo(curDate) == 0)
                                    {
                                        MessageBox.Show("Ревизия на указанную дату уже есть!");
                                        isDateFound = true;
                                        RefreshTable();
                                        break;
                                    }
                                    curRow++;
                                    curDate = DateTime.Parse(dataGridView1[0, dataGridView1.RowCount - curRow].Value.ToString());
                                }
                            }
                            catch (System.Exception ex)
                            {
                            	
                            }
                            

                            if(!isDateFound)
                            CommandExecution("INSERT INTO " + listView1.SelectedItems[0].Tag + " ([Дата]) VALUES (\"" + dataGridView1[0, e.RowIndex].Value + "\")");
                        
                        }
                    }


                    if(!isDateFound)
                    CommandExecution("UPDATE " + listView1.SelectedItems[0].Tag + " SET [" + dataGridView1.Columns[e.ColumnIndex].HeaderText + "]=\"" + editableText + "\" WHERE Код=" + dataGridView1[idCol, e.RowIndex].Value);
                }
                else
                {
                    RefreshTable();
                }

            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (listView1.SelectedItems.Count == 0)
                {
                    dataGridView1.Columns.Clear();
                    foreach (TextBoxWithColumnIndex tbwci in resultTextBoxes)
                    {
                        tbwci.tb.Dispose();
                    }
                }
                else
                {
                    dataGridView1.Columns.Clear();
                    RefreshTable();
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Убедитесь, что таблица \"" + listView1.SelectedItems[0].Text + "\" существует.\n\n\n" + ex.Message, "Ошибка обращения к таблице", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Width = this.Width - dataGridView1.Left - 25;
            dataGridView1.Height = this.Height - 90;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns[0].Visible = false;

            StreamReader sr = new StreamReader("shop.conf");
            string f = sr.ReadToEnd();
            this.Text = f + " - " + this.Text;
            sr.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text == passwordAdmin)
            {
                time.Enabled = true;
                textBox1.BackColor = Color.GreenYellow;
                passwordEntered = true;
            }
            else
            {
                time.Enabled = false;
                textBox1.BackColor = Color.White;
                passwordEntered = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int codeCol = 0;
            if (listView1.SelectedItems[0].Text == "Итоги")
                codeCol = 1;
            if (CheckRights())
                CommandExecution("DELETE FROM " + listView1.SelectedItems[0].Tag + " WHERE Код=" + dataGridView1[codeCol, dataGridView1.SelectedCells[0].RowIndex].Value);
        }

        void DGVAddColumn(DataGridView dgv, int displayIndex, string name)
        {
            /*DataGridViewColumn dgvc = new DataGridViewColumn();
            dgvc.DisplayIndex = displayIndex;
            dgvc.Name = name;
            //dgv.Columns.Insert(displayIndex, dgvc);
            dgv.Columns.Add(dgvc);
            //dgv.colum*/
            dgv.Columns.Add(name, name);
            dgv.Columns[name].DisplayIndex = displayIndex;
        }

        string ReadQuery(string file)
        {
            string resultQuery = "";
            try
            {
                StreamReader sr = new StreamReader(file);
                resultQuery = sr.ReadToEnd();
                sr.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Не могу прочитать файл!\n" + ex.Message);
            }
            return resultQuery;
        }

        void FillResult()
        {
            DataGridView dgv = new DataGridView();
            dgv.Visible = false;
            this.Controls.Add(dgv);

            command = conect.CreateCommand();

            command.CommandText = "SELECT Выручка.Дата FROM Выручка UNION SELECT Наценка.Дата FROM Наценка UNION SELECT ПриходТовара.Дата FROM ПриходТовара UNION SELECT ПриходФруктов.Дата FROM ПриходФруктов UNION SELECT Списание.Дата FROM Списание UNION SELECT СписаниеНаСклад.Дата FROM СписаниеНаСклад UNION SELECT СписаниеТары.Дата FROM СписаниеТары UNION SELECT СписаниеФрукты.Дата FROM СписаниеФрукты UNION SELECT Уценка.Дата FROM Уценка UNION SELECT УценкаФрукты.Дата FROM УценкаФрукты";
            command.CommandType = CommandType.Text;
            try
            {
                command.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
                command.CommandText = command.CommandText.Replace("\"\"", "0");
                command.ExecuteNonQuery();
            }

            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            DataTable table = new DataTable("Даты");

            adapter.Fill(table);
            dgv.DataSource = table.DefaultView;
            adapter.Update(table);

            dataGridView1.Rows.Add(dgv.RowCount);

            for (int i = 0; i < dgv.RowCount - 1; i++)
            {
                dataGridView1[1, i].Value = dgv[0, i].Value;
            }

            DGVAddColumn(dataGridView1, 4, "Сумма поступления Товара");
            DGVAddColumn(dataGridView1, 5, "Сумма поступления Фруктов");
            DGVAddColumn(dataGridView1, 6, "Сумма наценки");
            DGVAddColumn(dataGridView1, 7, "Сумма уценки");
            DGVAddColumn(dataGridView1, 8, "Сумма списания");
            DGVAddColumn(dataGridView1, 9, "Списание на склад");
            DGVAddColumn(dataGridView1, 10, "Уценка фруктов, овощей");
            DGVAddColumn(dataGridView1, 11, "Списание фруктов, овощей");
            DGVAddColumn(dataGridView1, 12, "Списание тара");
            DGVAddColumn(dataGridView1, 13, "Выручка нал");
            DGVAddColumn(dataGridView1, 14, "Выручка безнал");
            DGVAddColumn(dataGridView1, 15, "Итого выручка");
            DGVAddColumn(dataGridView1, 18, "Итог");

            for (int i = 8; i <= dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].ReadOnly = true;
                dataGridView1.Columns[i].DefaultCellStyle.BackColor = Color.LightGray;
            }

            for (int i = 0; i < dgv.RowCount - 1; i++)
            {
                DataGridView dgv2 = new DataGridView();
                dgv2.Visible = false;
                dgv2.Top = listView1.Top + listView1.Height;
                this.Controls.Add(dgv2);
                command = conect.CreateCommand();

                string sqlDate = dgv[0, i].Value.ToString().Remove(8).Replace('.', '/');
                string tmpDay = sqlDate.Remove(2);
                string tmpMonth = sqlDate.Substring(3).Remove(2);
                sqlDate = tmpMonth + "/" + tmpDay + sqlDate.Substring(5);

                command.CommandText = "SELECT Сумма FROM ПриходТовара WHERE Дата=#" + sqlDate + "#";
                command.CommandType = CommandType.Text;
                command.ExecuteNonQuery();

                OleDbDataAdapter dgv2adapter = new OleDbDataAdapter(command);
                DataTable dgv2table = new DataTable("Суммы");

                dgv2adapter.Fill(dgv2table);
                dgv2.DataSource = dgv2table.DefaultView;
                dgv2adapter.Update(dgv2table);
                dgv2.Dispose();
            }

            dgv.Dispose();
        }

        bool CheckRights()
        {

            if (passwordEntered)
                return true;

            if (dataGridView1.SelectedCells[0].OwningColumn.HeaderText == "Проверил")
            {
                MessageBox.Show("Недостаточно прав!");
                return false;
            }

            if (listView1.SelectedItems[0].Text == "Итоги")
            {
                MessageBox.Show("Недостаточно прав");
                return false;
            }

            int provIndex = FindColumn("Проверил");
            if ((provIndex > -1) && (dataGridView1[provIndex, dataGridView1.SelectedCells[0].RowIndex].Value.ToString() != ""))
            {
                MessageBox.Show("Данная позиция уже была проверена!");
                return false;
            }

            try
            {
                if ((dataGridView1[1, dataGridView1.SelectedCells[0].RowIndex].Value.ToString() == null)
                    || (dataGridView1[1, dataGridView1.SelectedCells[0].RowIndex].Value.ToString() == DateTime.Today.ToString())
                    || (dataGridView1[1, dataGridView1.SelectedCells[0].RowIndex].Value.ToString() == ""))
                    return true;
            }
            catch (System.Exception ex)
            {

            }

            MessageBox.Show("Недостаточно прав!");
            return false;

        }

        private void dataGridView1_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            foreach (TextBoxWithColumnIndex tbwci in resultTextBoxes)
            {
                RefreshTextBox(tbwci);
            }
        }

        void RefreshTextBox(TextBoxWithColumnIndex tbwci)
        {
            tbwci.tb.Left = GetColumnCoord(tbwci.columnIndex);
            tbwci.tb.Width = dataGridView1.Columns[tbwci.columnIndex].Width;
        }

        int GetColumnCoord(int columnIndex)
        {
            int leftCoord = 42;

            for (int i = 1; i < columnIndex; i++)
            {
                leftCoord += dataGridView1.Columns[i].Width;
            }
            leftCoord += dataGridView1.Left;

            return leftCoord;
        }

        private void dataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
            {
                int dev = e.NewValue - e.OldValue;

                foreach (TextBoxWithColumnIndex tbwci in resultTextBoxes)
                {
                    tbwci.tb.Left -= dev;
                }
            }
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //MessageBox.Show(dataGridView1.Columns[e.ColumnIndex].ValueType.Name);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            RefreshTable();
            if (dateTimePicker1.Value.CompareTo(dateTimePicker2.Value) == 1)
                dateTimePicker2.Value = dateTimePicker1.Value;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            RefreshTable();
            if (dateTimePicker1.Value.CompareTo(dateTimePicker2.Value) == 1)
                dateTimePicker1.Value = dateTimePicker2.Value;
        }
    }
}
