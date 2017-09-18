using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DWRReport
{
    public partial class DWRForm1 : Form
    {
        DataTable dt;
        DataRow dr;
        DataColumn idColumn;
        DataColumn nameColumn;
        BindingSource bindingSource1;
        public DWRForm1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataSet ds = new DataSet("DWR");
            bindingSource1 = new BindingSource();

            dt = new DataTable();
            idColumn = new DataColumn("ID", Type.GetType("System.Int32"));
            nameColumn = new DataColumn("Name", Type.GetType("System.String"));

            dt.Columns.Add(idColumn);
            dt.Columns.Add(nameColumn);

            dr = dt.NewRow();
            dr["ID"] = 1;
            dr["Name"] = "Name1";
            dt.Rows.Add(dr);

            ds.Tables.Add(dt);

            bindingSource1.DataSource = ds;

            dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
            dataGridView1.DataSource = bindingSource1;
        }
    }
}
