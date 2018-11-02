using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Windows;

namespace DB
{
    public partial class Form1 : Form
    {
        const string NameTitle = "DBEditor";
        private OleDbConnection connect = new OleDbConnection();
        private DataSet ds;
        private OleDbDataAdapter[] adapters = new OleDbDataAdapter[0];
        private DataGridView[] dataGrids = new DataGridView[0];
        private TabPage[] pages = new TabPage[0];


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Text = NameTitle;
            closeToolStripMenuItem.PerformClick();
        }

        private void newMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void openMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Microsoft Access (*.mdb; *.accdb)|*.mdb; *.accdb;";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                closeToolStripMenuItem.PerformClick();
                string path = openFileDialog1.FileName, file = openFileDialog1.SafeFileName;
                Text = file + " : Databse- " + path + $" (File format: .{file.Split('.')[1]}) - {NameTitle}";
                connect.ConnectionString = $@"Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source={path}";
                connect.Open();
                toolStripStatusLabel1.Text = "Reading Database";
                object[] r = new object[] { null, null, null, "TABLE" };
                DataTable dataTable = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, r);
                
                treeView1.Nodes.Clear();
                TreeNode[] newtree = new TreeNode[dataTable.Rows.Count];
                ds = new DataSet();
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    newtree[i] = new TreeNode(dataTable.Rows[i]["TABLE_NAME"].ToString());
                    Array.Resize(ref adapters, adapters.Length + 1);
                    adapters[i] = new OleDbDataAdapter(new OleDbCommand($"SELECT * FROM {newtree[i].Text}", connect));
                    adapters[i].Fill(ds, newtree[i].Text);
                    Array.Resize(ref dataGrids, dataGrids.Length + 1);
                    dataGrids[i] = new DataGridView{ DataSource = ds.Tables[newtree[i].Text] };
                }
                treeView1.Nodes.AddRange(new TreeNode[] { new TreeNode("Tables", newtree) });
                treeView1.Visible = true;
                connect.Close();
                toolStripStatusLabel1.Text = "Ready";
            }
        }

        private void saveMenuItem_Click(object sender, EventArgs e)
        {
            //for (int i = 0; i < adapters.Length; i++)
            //{
            //    adapters[i].InsertCommand(new OleDbCommand($"UPDATE *"))
            //    adapters[i].Update(ds.Tables[i]);
            //}
        }

        private void saveAsMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }

        private void treeView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode sel = treeView1.HitTest(e.Location).Node;
            if (sel != null && connect.ConnectionString != "")
            {
                toolStripStatusLabel1.Text = "Datasheet View";
                if (!tabControl1.Contains(tabControl1.TabPages[sel.Text]))
                {
                    tabControl1.TabPages.Add(sel.Text, sel.Text);
                    //dataGrids[sel.Index] = new DataGridView { DataSource = ds.Tables[sel.Text] };
                    dataGrids[sel.Index].Parent = tabControl1.TabPages[sel.Text];
                    dataGrids[sel.Index].Dock = DockStyle.Fill;
                    dataGrids[sel.Index].BorderStyle = BorderStyle.None;
                    dataGrids[sel.Index].EditMode = DataGridViewEditMode.EditProgrammatically;
                    tabControl1.TabPages[sel.Text].ContextMenuStrip = contextMenuStrip1;
                }
                tabControl1.SelectTab(sel.Text);
                //dataGridView1.DataSource = ds.Tables[sel.Text];
            }
        }

        private void closeMenuItem_Click(object sender, EventArgs e)
        {
            connect.ConnectionString = "";
            if(ds != null)
                ds.Dispose();
            Array.Resize(ref adapters, 0);
            //dataGridView1.DataSource = null;
            treeView1.Visible = false;
            Text = NameTitle;
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabControl1.SelectedTab);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Clear();
        }
    }
}
