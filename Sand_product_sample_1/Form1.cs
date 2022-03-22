using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using Spire.Xls;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using MetroFramework;
using MetroFramework.Forms;


namespace Sand_product_sample_1
{
    public partial class Frontend : MetroForm
    {
        //SqlConnection connection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Moureen\Desktop\appdev\desktop_apps\ghf\Sand_product_sample_1\Sand_product_sample_1\Sand_product_sample_1_db.mdf;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");
        SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["Sand_product_sample_1_dbConString"].ConnectionString);
        public Exception arg;
        Excel.Application xlApp = new Excel.Application();
        string programpath = Application.StartupPath;
        object misValue = System.Reflection.Missing.Value;
        Excel.Workbook workbook1;

        public Frontend()
        {
            InitializeComponent();
            ConnectionCheck();
            SqlCommand mycmd = new SqlCommand("SET DATEFORMAT dmy;", connection);
            mycmd.ExecuteNonQuery();
            ConnectionCheck();
            table_relations();
            expensetaboptions.DrawItem += new DrawItemEventHandler(expensetaboptions_DrawItem);
            newexpensedateTimePicker.Value = DateTime.Now;
            newproductdateTimePicker.Value = DateTime.Now;
            productmiscellaneousdateTimePicker.Value = DateTime.Now;
            productdamagedateTimePicker.Value = DateTime.Now;
            creditdepositdateTimePicker.Value = DateTime.Now;
            rawmaterialpurchasedate.Value = DateTime.Now;
            rawmaterialdiscarddateTime.Value = DateTime.Now;
            moulddateTimePicker.Value = DateTime.Now;
        }
        private void ConnectionCheck()
        {
            connectionbutton.FlatStyle = FlatStyle.Flat;
            connectionbutton.FlatAppearance.BorderSize = 0;
            try
            {
                connection.Close();
                connection.Open();
                connectionbutton.BackColor = Color.DarkTurquoise;
            }
            catch (Exception ex)
            {
                connectionbutton.BackColor = Color.Red;
                arg = ex;
            }
        }
        private void Clearcontrols()
        {
            //input expense controls
            newexpensedateTimePicker.Value = DateTime.Now;
            expensedetailstextBox.Text = "";
            expenseamountextBox.Text = "";
            expensetypecomboBox.SelectedIndex = 0;
            expnsenamecomboBox.SelectedIndex = 0;

            //new product receipt controls
            productlistcombobox.Items.Clear();
            newproductdateTimePicker.Value = DateTime.Now;
            //otherproductnametextBox.Text = "";
            unitmeasurementtextBox.Text = "";
            newproductreceiptdesctextbox.Text = "";
            unitmeasurementtextBox.Text = "";
            unitmeasurementtextBox.ReadOnly = true;
            expensedetailstextBox.Text = "";
            expenseamountextBox.Text = "";
            editunitpricecheckBox.Checked = false;
            editunitpricecheckBox.Enabled = false;
            editunitpricetextBox.ReadOnly = true;
            editunitpricetextBox.Text = "";
            unitpricetextBox.Text = "";
            unitpricetextBox.ReadOnly = true;
            unitcosttextBox.Text = "";
            finequantityreceipttextBox.Text = "";
            damagequantityreceipttextBox.Text = "";

            //product sale controls
            productsalecomboBox.Items.Clear();
            newproductsaledateTimePicker.Value = DateTime.Now;
            voucher_idtextBox.Text = "";
            finequantitysoldtextBox.Text = "0";
            fineunitpricetextBox.Text = "0";
            damagequantitysoldtextBox.Text = "0";
            damageunitpricetextBox.Text = "0";
            customernametextBox.Text = "";
            totalsaleamounttextBox.Text = "";
            
            //productmiscellaneous controls
            productmiscellaneousdateTimePicker.Value = DateTime.Now;
            productmiscellaneousfinequantitytextBox.Text = "0";
            productmiscellaneousdamagequantitytextBox.Text = "0";
            productmiscellaneousreasontextBox.Text = "";
            productmiscellaneousunitpricetextBox.Text = "0";
            productmiscellaneoustotalvaluetextBox.Text = "0";

            //product damage controls
            productdamagedateTimePicker.Value = DateTime.Now;
            productdamagequantitytextBox.Text = "0";

            //credit deposit controls
            creditcustomernamecomboBox.Items.Clear();
            creditdepositdateTimePicker.Value = DateTime.Now;
            receiptidtextBox.Text = "";
            creditcurrentamounttextBox.Text = "0";
            creditpaymentamounttextBox.Text = "0";

            //raw material purchase controls
            rawmaterialpurchasedate.Value = DateTime.Now;
            otherrawmaterialpurchasenametextBox.Text = "";
            otherrawmaterialpurchasenametextBox.ReadOnly = true;
            rawmaterialpurchaseunitofmeasurementtextBox.Text = "";
            rawmaterialpurchaseunitofmeasurementtextBox.ReadOnly = true;
            rawmaterialpurchaseunitcosttextBox.ReadOnly = true;
            rawmaterialpurchaseunitcosttextBox.Text = "0";
            editrawmaterialunitcostcheckBox.Enabled = false;
            editrawmaterialunitcostcheckBox.Checked = false;
            rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
            rawmaterialpurchaseeditunitcosttextBox.Text = "0";
            rawmaterialpurchasequantityreceivedtextBox.Text = "0";
            rawmaterialpurchasedesctextBox.Text = "0";

            //raw material discard controls
            rawmaterialdiscarddateTime.Value = DateTime.Now;
            rawmaterialquantitydiscardedtextBox.Text = "0";
            rawmaterialdiscardreasontextBox.Text = "";
            rawmaterialunitcosttextBox.Text = "0";

            //raw material product allocation controls
            rmproductallocationproductnamecomboBox.Items.Clear();
            rmproductallocationunittextBox.Text = "0";
            rmproductallocationunitcosttextBox.Text = "0";
            newrmproductallocationproductnametextBox.Text = "";
            newrmproductallocationproductnametextBox.ReadOnly = true;

            //Mould Controls
            moulddateTimePicker.Value = DateTime.Now;
            mouldnametextBox.Text = "";
            moulddepartmenttextBox.Text = "";
            mouldproductnametextBox.Text = "";
            mouldquantitytextBox.Text = "0";
            mouldratetextBox.Text = "0";
            mouldpricetextBox.Text = "0";
            mouldreceivedtextBox.Text = "0";
            mouldratetotaltextBox.Text = "";
            mouldreceivedtotaltextBox.Text = "";


            rawmaterialpurchasenamecomboBox.Items.Clear();
            Form1_Loader();
           
            productlistcombobox.SelectedIndex = 0;
            rawmaterialpurchasenamecomboBox.SelectedIndex = 0;
            rmproductallocationproductnamecomboBox.SelectedIndex = 0;

            this.raw_material_tableTableAdapter.Fill(this.rawmaterialnameDataSet.raw_material_table);
            this.product_tableTableAdapter.Fill(this.productnamelistDataset.Product_table);
            this.expenses_tableTableAdapter.Fill(this.expenseDataSet.expenses_table);
            this.mould_tableTableAdapter.Fill(this.mould_dbDataSet.Mould_table);
        }
        private void Form1_Loader()
        {
            // TODO: This line of code loads data into the 'rawmaterialnameDataSet.raw_material_table' table. You can move, or remove it, as needed.
            this.raw_material_tableTableAdapter.Fill(this.rawmaterialnameDataSet.raw_material_table);
            // TODO: This line of code loads data into the 'productnamelistDataset.Product_table' table. You can move, or remove it, as needed.
            this.product_tableTableAdapter.Fill(this.productnamelistDataset.Product_table);
            // TODO: This line of code loads data into the 'sand_product_sample_1_dbDataSet.expenses_table' table. You can move, or remove it, as needed.
            this.expenses_tableTableAdapter.Fill(this.expenseDataSet.expenses_table);
            // TODO: This line of code loads data into the 'mould_dbDataSet.Mould_table' table. You can move, or remove it, as needed.
            this.mould_tableTableAdapter.Fill(this.mould_dbDataSet.Mould_table);
            table_relations();
            ConnectionCheck();
            SqlCommand command2 = new SqlCommand("Select sum(total_credit_amount) total_credit_amount, sum(current_credit_amount) current_credit_amount from product_customer_credit_table;", connection);
            SqlDataReader reader2 = command2.ExecuteReader();
            while (reader2.Read())
            {
                totalcreditamountlabellabel.Text = reader2["total_credit_amount"].ToString();
                currentcreditamountlabel.Text = reader2["current_credit_amount"].ToString();
            }

            ConnectionCheck();
            productlistcombobox.Items.Add("");
            SqlCommand cmd = new SqlCommand("Select product_name from raw_m_product_temp_table", connection);
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                string productname = rdr["product_name"].ToString();
                productlistcombobox.Items.Add(productname);
            }
            ConnectionCheck();
            SqlCommand cmd2 = new SqlCommand("Select product_name from product_table where total_quantity_available>0", connection);
            SqlDataReader rdr2 = cmd2.ExecuteReader();
            while (rdr2.Read())
            {
                string product_name = rdr2["product_name"].ToString();
                productsalecomboBox.Items.Add(product_name);
            }
            ConnectionCheck();
            SqlCommand cmd3 = new SqlCommand("Select customer_name from product_customer_credit_table where total_credit_amount != amount_paid", connection);
            SqlDataReader rdr3 = cmd3.ExecuteReader();
            while (rdr3.Read())
            {
                creditcustomernamecomboBox.Items.Add(rdr3["customer_name"].ToString());

            }
            ConnectionCheck();
            rawmaterialpurchasenamecomboBox.Items.Add("");
            rawmaterialpurchasenamecomboBox.Items.Add("Other");
            SqlCommand cmd5 = new SqlCommand("Select raw_material_name from raw_material_table;", connection);
            SqlDataReader rdr5 = cmd5.ExecuteReader();
            while (rdr5.Read())
            {
                rawmaterialpurchasenamecomboBox.Items.Add(rdr5["raw_material_name"].ToString());
                
            }

            ConnectionCheck();
            rmproductallocationproductnamecomboBox.Items.Add("");
            rmproductallocationproductnamecomboBox.Items.Add("Other");
            SqlCommand cmd6 = new SqlCommand("Select product_name from raw_m_product_temp_table", connection);
            SqlDataReader rdr6 = cmd6.ExecuteReader();
            while (rdr6.Read())
            {
                rmproductallocationproductnamecomboBox.Items.Add(rdr6["product_name"].ToString());
            }
            

        }
        private void new_expense_panel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Form1_Loader();
        }

        private void connectionbutton_Click(object sender, EventArgs e)
        {
            if (connectionbutton.BackColor.Equals(Color.Red))
            {
                MessageBox.Show("There is a problem with the conncection:  " + arg, "Error");
            }
        }


        private void expensetaboptions_DrawItem(object sender, DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            Brush _textBrush;

            // Get the item from the collection.
            TabPage _tabPage = expensetaboptions.TabPages[e.Index];

            // Get the real bounds for the tab rectangle.
            Rectangle _tabBounds = expensetaboptions.GetTabRect(e.Index);

            if (e.State == DrawItemState.Selected)
            {
                // Draw a different background color, and don't paint a focus rectangle.
                _textBrush = new SolidBrush(Color.Black);
                g.FillRectangle(Brushes.LightGray, e.Bounds);
            }
            else
            {
                _textBrush = new System.Drawing.SolidBrush(e.ForeColor);
                e.DrawBackground();
            }

            // Use our own font.
            Font _tabFont = new Font("Maindra CD", 11.25f, FontStyle.Regular, GraphicsUnit.Pixel);

            // Draw string. Center the text.
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }
        private void producttaboptions_DrawItem(object sender, DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            Brush _textBrush;

            // Get the item from the collection.
            TabPage _tabPage = producttaboptions.TabPages[e.Index];

            // Get the real bounds for the tab rectangle.
            Rectangle _tabBounds = producttaboptions.GetTabRect(e.Index);

            if (e.State == DrawItemState.Selected)
            {
                // Draw a different background color, and don't paint a focus rectangle.
                _textBrush = new SolidBrush(Color.Black);
                g.FillRectangle(Brushes.LightGray, e.Bounds);
            }
            else
            {
                _textBrush = new System.Drawing.SolidBrush(e.ForeColor);
                e.DrawBackground();
            }

            // Use our own font.
            Font _tabFont = new Font("Maindra CD", 11.25f, FontStyle.Regular, GraphicsUnit.Pixel);

            // Draw string. Center the text.
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }

        private void rawmaterialoptionstab_DrawItem(object sender, DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            Brush _textBrush;

            // Get the item from the collection.
            TabPage _tabPage = rawmaterialoptionstab.TabPages[e.Index];

            // Get the real bounds for the tab rectangle.
            Rectangle _tabBounds = rawmaterialoptionstab.GetTabRect(e.Index);

            if (e.State == DrawItemState.Selected)
            {
                // Draw a different background color, and don't paint a focus rectangle.
                _textBrush = new SolidBrush(Color.Black);
                g.FillRectangle(Brushes.LightGray, e.Bounds);
            }
            else
            {
                _textBrush = new System.Drawing.SolidBrush(e.ForeColor);
                e.DrawBackground();
            }

            // Use our own font.
            Font _tabFont = new Font("Maindra CD", 11.25f, FontStyle.Regular, GraphicsUnit.Pixel);

            // Draw string. Center the text.
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }

        private void applicationexitbutton_Click(object sender, EventArgs e)
        {
            table_relations();
            connection.Close();
            System.Windows.Forms.Application.Exit();
        }

        private void expensesearchbutton_Click(object sender, EventArgs e)
        {
            table_relations();
            connection.Close();
            connection.Open();
            string expensesearchquery = "SET DATEFORMAT dmy; Select  date, expense_name, expense_type, expense_detail,amount from expenses_table where format(date,'dd-MMM-yyyy') like '%" + searchexpensetextBox.Text + "%' or expense_name like '%" + searchexpensetextBox.Text + "%' or expense_type like '%" + searchexpensetextBox.Text + "%' or expense_detail like '%" + searchexpensetextBox.Text + "%'or amount like '%" + searchexpensetextBox.Text + "%'";
            SqlCommand com = new SqlCommand(expensesearchquery, connection);
            SqlDataAdapter data_adapter = new SqlDataAdapter(com);
            DataTable data_table = new DataTable();
            data_adapter.Fill(data_table);
            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = data_table;
            searchexpensedataGridView.DataSource = bindingSource;
            data_adapter.Update(data_table);
            SqlDataReader reader;
            reader = com.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();
                searchexpensedataGridView.Columns[0].HeaderText = "Date";
                searchexpensedataGridView.Columns[0].Width = 140;
                searchexpensedataGridView.Columns[1].HeaderText = "Expense Name";
                searchexpensedataGridView.Columns[1].Width = 140;
                searchexpensedataGridView.Columns[2].HeaderText = "Expense Type";
                searchexpensedataGridView.Columns[2].Width = 140;
                searchexpensedataGridView.Columns[3].HeaderText = "Expense Detail";
                searchexpensedataGridView.Columns[3].Width = 140;
                searchexpensedataGridView.Columns[4].HeaderText = "Amount";
                searchexpensedataGridView.Columns[4].Width = 140;
            }
        }

        private void expensestableBindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void expenseinsertbutton_Click(object sender, EventArgs e)
        {
            if (expnsenamecomboBox.SelectedItem == null || expnsenamecomboBox.SelectedIndex == 0 || expensetypecomboBox.SelectedItem == null || expensetypecomboBox.SelectedIndex == 0 || String.IsNullOrEmpty(expensedetailstextBox.Text) || String.IsNullOrEmpty(expenseamountextBox.Text))
            {

                MessageBox.Show("Please fill the all the fields", "Notice");

            }
            else
            {
                ConnectionCheck();
                string expenseinsertsearchquery = "SET DATEFORMAT dmy; Insert into expenses_table (date, expense_name, expense_type, expense_detail, amount) VALUES('" + newexpensedateTimePicker.Value.ToShortDateString() + "', '" + expnsenamecomboBox.SelectedItem.ToString() + "', '" + expensetypecomboBox.SelectedItem.ToString() + "', '" + expensedetailstextBox.Text + "', '" + expenseamountextBox.Text + "')";
                SqlCommand com = new SqlCommand(expenseinsertsearchquery, connection);
                try
                {

                    com.ExecuteNonQuery();
                    MessageBox.Show("Entry successfully inserted. ", "Report", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    table_relations();
                    expensefieldsclearbutton_Click(sender, e);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n\n" + ex, "Error");
                }
            }
            this.expenses_tableTableAdapter.Fill(this.expenseDataSet.expenses_table);
        }

        private void expensefieldsclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }
        private void expenseamountextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(expenseamountextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                expenseamountextBox.Text = numberOnly;
                expenseamountextBox.SelectionStart = expenseamountextBox.Text.Length;
            }
            else
            {
                expenseamountextBox.Text = "";
            }
        }


        private void productlistcombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (productlistcombobox.SelectedItem.ToString() == "Other")
            {
                //otherproductnametextBox.ReadOnly = false;
                //otherproductnametextBox.Text = "";
                unitpricetextBox.ReadOnly = false;
                editunitpricecheckBox.Enabled = false;
                editunitpricecheckBox.Checked = false;
                unitpricetextBox.Text = "";
                unitmeasurementtextBox.ReadOnly = false;
                unitmeasurementtextBox.Text = "";
            }
            else if (productlistcombobox.SelectedItem.ToString() == "")
            {
                unitpricetextBox.Text = "0";
                unitpricetextBox.ReadOnly = true;
                editunitpricetextBox.Text = "";
                editunitpricetextBox.ReadOnly = true;
                //otherproductnametextBox.ReadOnly = true;
                //otherproductnametextBox.Text = "";
                editunitpricecheckBox.Enabled = false;
                editunitpricecheckBox.Checked = false;
                unitmeasurementtextBox.ReadOnly = true;
                unitmeasurementtextBox.Text = "";
            }
            else
            {
                ConnectionCheck();
                //otherproductnametextBox.ReadOnly = true;
                //otherproductnametextBox.Text = "";
                SqlCommand cmd = new SqlCommand("Select current_unit_price,unit_of_measurement from product_table where product_name = '" + productlistcombobox.SelectedItem.ToString() + "'", connection);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    unitpricetextBox.Text = rdr["current_unit_price"].ToString();
                    unitmeasurementtextBox.Text = rdr["unit_of_measurement"].ToString();
                    unitmeasurementtextBox.ReadOnly = true;
                    unitpricetextBox.ReadOnly = true;
                    editunitpricecheckBox.Checked = false;
                    ConnectionCheck();
                    SqlCommand unitcostquery = new SqlCommand("SET DATEFORMAT dmy; select A.product_id ,sum(A.product_unit_cost) unit_cost from (Select A.rm_id, A.product_id product_id, (A.rm_unit_quantity_utilized*B.unit_cost) product_unit_cost from Raw_matrial_product_allocation_table as A, Raw_material_purchase_table as B where A.rm_id = B.rm_id)A inner join Product_table AS B on upper(A.product_name) = upper(B.product_name) where upper(B.product_name) = upper('" + productlistcombobox.SelectedItem.ToString() + "') group by A.product_id;", connection);
                    SqlDataReader rdr2 = unitcostquery.ExecuteReader();
                    if (rdr2.Read())
                    { unitcosttextBox.Text = rdr2["unit_cost"].ToString(); }
                    else { unitcosttextBox.Text = "0"; }
                    if (String.IsNullOrEmpty(unitmeasurementtextBox.Text) || String.IsNullOrEmpty(unitpricetextBox.Text))
                    {
                        unitmeasurementtextBox.ReadOnly = false;
                        unitpricetextBox.ReadOnly = false;
                    }
                }

            }
        }
        private void newproductreceiptbuttonbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void newproductreceiptpanel_MouseHover(object sender, EventArgs e)
        {
            if (/*otherproductnametextBox.ReadOnly == true &&*/ unitpricetextBox.ReadOnly == true && unitpricetextBox.Text.ToString() != "")
            {
                editunitpricecheckBox.Enabled = true;
                editunitpricetextBox.Text = "";
            }
            else
            {
                editunitpricecheckBox.Enabled = false;
                editunitpricecheckBox.Checked = false;
            }
            if (editunitpricecheckBox.Checked == true)
            {
                editunitpricetextBox.ReadOnly = false;
                //editunitpricetextBox.Text = "";
            }
            else
            {
                editunitpricetextBox.ReadOnly = true;
                editunitpricetextBox.Text = "";
            }
        }

        private void unitpricetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(unitpricetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                unitpricetextBox.Text = numberOnly;
                unitpricetextBox.SelectionStart = unitpricetextBox.Text.Length;
            }
            else
            {
                unitpricetextBox.Text = "";
            }
        }
        private void editunitpricetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(editunitpricetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                editunitpricetextBox.Text = numberOnly;
                editunitpricetextBox.SelectionStart = editunitpricetextBox.Text.Length;
            }
            else
            {
                editunitpricetextBox.Text = "";
            }
        }
        private void rawmaterialpurchasequantityreceivedtextBox_TextChanged_1(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialpurchasequantityreceivedtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = numberOnly;
                rawmaterialpurchasequantityreceivedtextBox.SelectionStart = rawmaterialpurchasequantityreceivedtextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = "";
            }
        }
        private void unitcosttextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(unitcosttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                unitcosttextBox.Text = numberOnly;
                unitcosttextBox.SelectionStart = unitcosttextBox.Text.Length;
            }
            else
            {
                unitcosttextBox.Text = "";
            }
        }

        private void finequantityreceipttextBox_TextChanged_1(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(finequantityreceipttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                finequantityreceipttextBox.Text = numberOnly;
                finequantityreceipttextBox.SelectionStart = finequantityreceipttextBox.Text.Length;
            }
            else
            {
                finequantityreceipttextBox.Text = "";
            }
        }

        private void damagequantityreceipttextBox_TextChanged_1(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(damagequantityreceipttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                damagequantityreceipttextBox.Text = numberOnly;
                damagequantityreceipttextBox.SelectionStart = damagequantityreceipttextBox.Text.Length;
            }
            else
            {
                damagequantityreceipttextBox.Text = "";
            }
        }


        private void newproductreceiptinsertbutton_Click(object sender, EventArgs e)
        {
            string producttablequery = "";
            if (unitcosttextBox.Text == "0" || productlistcombobox.SelectedItem == null || productlistcombobox.SelectedIndex == 0 || String.IsNullOrEmpty(unitmeasurementtextBox.Text) || String.IsNullOrEmpty(unitpricetextBox.Text) || String.IsNullOrEmpty(finequantityreceipttextBox.Text) || String.IsNullOrEmpty(damagequantityreceipttextBox.Text))
            {

                MessageBox.Show("Please fill the all the fields", "Notice");

            }
            else if (productlistcombobox.SelectedItem.ToString() == "Other" /*&& String.IsNullOrEmpty(otherproductnametextBox.Text)*/)
            {
                MessageBox.Show("Please fill the all the fields", "Notice");
            }
            else if (editunitpricecheckBox.Checked == true && String.IsNullOrEmpty(editunitpricetextBox.Text))
            {
                MessageBox.Show("Please fill the all the fields", "Notice");
            }
            else
            {
                
                int counta = 0;
                SqlCommand cmd12 = new SqlCommand("SET DATEFORMAT dmy; Select count(*) counta from product_table where upper(product_name) = upper('" + productlistcombobox.SelectedItem.ToString() + "')", connection);
                SqlDataReader rdr12 = cmd12.ExecuteReader();
                while (rdr12.Read())
                {
                    counta = Convert.ToInt32(rdr12["counta"]);
                }
                if (counta == 0)
                {
                    producttablequery = "SET DATEFORMAT dmy; Insert into product_table (date_added,product_name, product_desc, unit_of_measurement, current_unit_price, current_unit_cost) VALUES('" + newproductdateTimePicker.Value.ToShortDateString() + "', '" + productlistcombobox.SelectedItem.ToString() + "','" + newproductreceiptdesctextbox.Text + "','" + unitmeasurementtextBox + "','" + unitpricetextBox.Text + "', '" + unitcosttextBox.Text + "');";
                    producttablequery += "Select MAX(product_id) product_id from product_table;";
                }
                else
                {
                    if (editunitpricecheckBox.Checked == true)
                    {
                        producttablequery = "SET DATEFORMAT dmy; Update product_table set current_unit_price = '" + editunitpricetextBox.Text + "', current_unit_cost = '" + unitcosttextBox.Text + "' where upper(product_name) = upper('" + productlistcombobox.SelectedItem.ToString() + "');";
                        producttablequery += "Select product_id from product_table where upper(product_name) = upper('" + productlistcombobox.SelectedItem.ToString() + "');";

                    }
                    else
                    {
                        producttablequery = "Select product_id from product_table where upper(product_name) = upper('" + productlistcombobox.SelectedItem.ToString() + "');";
                    }
                }
                int productid = 0;
                SqlCommand query_table = new SqlCommand(producttablequery, connection);
                try
                {
                    ConnectionCheck();
                    productid = (Int32)query_table.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update Product table! " + ex, "Error");
                }
                string productreceivedtablequery = "SET DATEFORMAT dmy; Insert into product_received_table (date_of_receipt, product_id, fine_quantity_received, damaged_quantity_received,unit_cost) VALUES ('" + newproductdateTimePicker.Value.ToShortDateString() + "','" + productid + "', '" + finequantityreceipttextBox.Text + "','" + damagequantityreceipttextBox.Text + "','" + unitcosttextBox.Text + "');";
                if (damagequantityreceipttextBox.Text != "" || damagequantityreceipttextBox.Text != "0")
                {
                    productreceivedtablequery += "SET DATEFORMAT dmy; Insert into product_damage_table (date, product_id, period_of_damage, quantity_damaged, unit_cost) VALUES('" + newproductdateTimePicker.Value.ToShortDateString() + "','" + productid + "','On Receipt','" + damagequantityreceipttextBox.Text + "', '" + unitcosttextBox.Text + "');";
                }
                if (productid != 0)
                {
                    SqlCommand com = new SqlCommand(productreceivedtablequery, connection);
                    try
                    {
                        ConnectionCheck();
                        com.ExecuteNonQuery();
                        string meee2;
                        int abc = 0;
                        ArrayList rmidlist = new ArrayList();
                        { meee2 = productlistcombobox.SelectedItem.ToString(); }
                        SqlCommand rmcmd5 = new SqlCommand("Select rm_id from raw_material_product_allocation_table where upper(product_name) = upper ('" + meee2 + "');", connection);
                        ConnectionCheck();
                        SqlDataReader rmrdr5 = rmcmd5.ExecuteReader();
                        while (rmrdr5.Read())
                        {
                            rmidlist.Add(rmrdr5["rm_id"].ToString());
                        }
                        abc = rmidlist.Count;
                        for (int d = 0; d < abc; d++)
                        {
                            string resultspool = "SET DATEFORMAT dmy; Select A.rm_id, A.rm_unit_quantity_utilized rm_unit_quantity_utilized , B.current_unit_cost current_unit_cost from raw_material_product_allocation_table AS A, raw_material_current_cost_table AS B, raw_material_table AS C where A.rm_id = C.rm_id and upper(B.raw_material_name) = upper(C.raw_material_name) and A.rm_id = '" + rmidlist[d] + "';";
                            SqlCommand resultcmd = new SqlCommand(resultspool, connection);
                            SqlDataReader resultrdr = resultcmd.ExecuteReader();
                            while (resultrdr.Read())
                            {
                                string rmusageinsert = "SET DATEFORMAT dmy; Insert into raw_material_to_product_usage_table (date_used, rm_id, quantity_used, rm_unit_cost, product_id) VALUES('" + newproductdateTimePicker.Value.ToString() + "', '" + rmidlist[d] + "','" + (Convert.ToDouble(resultrdr["rm_unit_quantity_utilized"].ToString()) * (Convert.ToDouble(finequantityreceipttextBox.Text) + (Convert.ToDouble(damagequantityreceipttextBox.Text)))) + "','" + resultrdr["current_unit_cost"].ToString() + "',(select produuct_id from product_table where upper(product_name) = upper('" + meee2 + "'));";

                                SqlCommand finacmd = new SqlCommand(rmusageinsert, connection);
                                ConnectionCheck();
                                finacmd.ExecuteNonQuery();
                            }
                        }

                        MessageBox.Show("Product updated successfully.", "Success");

                        table_relations();
                        Clearcontrols();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failure to update Product received table! " + ex, "Error");
                    }

                }
            }
        }


        private void fineunitpricetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(fineunitpricetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                fineunitpricetextBox.Text = numberOnly;
                fineunitpricetextBox.SelectionStart = fineunitpricetextBox.Text.Length;
            }
            else
            {
                fineunitpricetextBox.Text = "";
            }
        }
        private void finequantitysoldtextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(finequantitysoldtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                finequantitysoldtextBox.Text = numberOnly;
                finequantitysoldtextBox.SelectionStart = finequantitysoldtextBox.Text.Length;
            }
            else
            {
                finequantitysoldtextBox.Text = "";
            }
        }
        private void damagequantitysoldtextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(damagequantitysoldtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                damagequantitysoldtextBox.Text = numberOnly;
                damagequantitysoldtextBox.SelectionStart = damagequantitysoldtextBox.Text.Length;
            }
            else
            {
                damagequantitysoldtextBox.Text = "";
            }
        }
        private void damageunitpricetextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            string numberOnly = Regex.Replace(damageunitpricetextBox.Text, "[^0-9.+-]", "");
            damageunitpricetextBox.Text = numberOnly;
            damageunitpricetextBox.SelectionStart = damageunitpricetextBox.Text.Length;
        }

        private void damageunitpricetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(damageunitpricetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                damageunitpricetextBox.Text = numberOnly;
                damageunitpricetextBox.SelectionStart = damageunitpricetextBox.Text.Length;
            }
            else
            {
                damageunitpricetextBox.Text = "";
            }
        }
        private void panel1_MouseHover(object sender, EventArgs e)
        {
            float num;
            if (!float.TryParse(finequantitysoldtextBox.Text, out num))
            { finequantitysoldtextBox.Text = "0"; }

            if (!float.TryParse(fineunitpricetextBox.Text, out num))
            { fineunitpricetextBox.Text = "0"; }

            if (!float.TryParse(damagequantitysoldtextBox.Text, out num))
            { damagequantitysoldtextBox.Text = "0"; }

            if (!float.TryParse(damageunitpricetextBox.Text, out num))
            { damageunitpricetextBox.Text = "0"; }
            totalsaleamounttextBox.Text = ((Convert.ToDouble(finequantitysoldtextBox.Text) * Convert.ToDouble(fineunitpricetextBox.Text)) + (Convert.ToDouble(damagequantitysoldtextBox.Text) * Convert.ToDouble(damageunitpricetextBox.Text))).ToString();
        }


        private void peoductsaleinsertbutton_Click(object sender, EventArgs e)
        {
            if (totalsaleamounttextBox.Text == "0" || productsalecomboBox.SelectedItem == null || String.IsNullOrEmpty(voucher_idtextBox.Text) || paymenttypecomboBox.SelectedItem == null || String.IsNullOrEmpty(customernametextBox.Text))
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else if ((finequantitysoldtextBox.Text == "0" && fineunitpricetextBox.Text != "0") || (finequantitysoldtextBox.Text != "0" && fineunitpricetextBox.Text == "0") || (damageunitpricetextBox.Text == "0" && damagequantitysoldtextBox.Text != "0") || (damageunitpricetextBox.Text != "0" && damagequantitysoldtextBox.Text == "0"))
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else
            {
                string salesinsertsearchquery = "SET DATEFORMAT dmy; Insert into product_sales_table (date_of_sale, product_id, voucher_id, payment_type, fine_quantity_sold, fine_unit_price, damages_quantity_sold, damaged_unit_price, customer_name, unit_cost) VALUES('" + newproductsaledateTimePicker.Value.ToShortDateString() + "', (Select product_id from product_table where upper(product_name) = upper('" + productsalecomboBox.SelectedItem.ToString() + "') ), '" + voucher_idtextBox.Text + "', '" + paymenttypecomboBox.SelectedItem.ToString() + "', '" + finequantitysoldtextBox.Text + "','" + fineunitpricetextBox.Text + "','" + damagequantitysoldtextBox.Text + "','" + damageunitpricetextBox.Text + "','" + customernametextBox.Text + "', '" + currentunitcosttextBox.Text + "');";

                if (paymenttypecomboBox.SelectedItem.ToString() == "Credit")
                {
                    SqlCommand cmd2 = new SqlCommand("SET DATEFORMAT dmy; Select count(*) from product_customer_credit_table where upper(customer_name) = upper('" + customernametextBox.Text + "'));", connection);
                    int product = 0;
                    product = (Int32)cmd2.ExecuteScalar();
                    if (product == 0)
                    {
                        salesinsertsearchquery += "SET DATEFORMAT dmy; Insert into product_customer_credit_table (date_of_sale, customer_name, total_credit_amount) VALUES ('" + newproductsaledateTimePicker.Value.ToShortDateString() + "',upper('" + customernametextBox.Text + "'),'" + totalsaleamounttextBox.Text + "');";
                    }
                    else
                    {
                        salesinsertsearchquery += "SET DATEFORMAT dmy; Update product_customer_credit_table set total_credit_amount = (total_credit_amount + " + Convert.ToDouble(totalsaleamounttextBox.Text) + ") where upper(customer_name) = upper('" + customernametextBox.Text + "');";
                    }

                    SqlCommand cmd3 = new SqlCommand("SET DATEFORMAT dmy; Select count(*) from product_credit_table where voucher_id = upper('" + voucher_idtextBox.Text + "'));", connection);
                    int product2 = 0;
                    ConnectionCheck();
                    product = (Int32)cmd3.ExecuteScalar();
                    if (product2 == 0)
                    {
                        salesinsertsearchquery += "SET DATEFORMAT dmy; Insert into product_credit_table (date_of_sale, voucher_id, customer_name, credit_amount) VALUES ('" + newproductsaledateTimePicker.Value.ToShortDateString() + "','" + voucher_idtextBox.Text + "','" + customernametextBox.Text + "','" + totalsaleamounttextBox.Text + "');";
                    }
                    else
                    {
                        salesinsertsearchquery += "SET DATEFORMAT dmy; Update product_credit_table set credit_amount = (credit_amount + " + Convert.ToDouble(totalsaleamounttextBox.Text) + ") where voucher_id = '" + voucher_idtextBox.Text + "';";
                    }
                }
                SqlCommand com = new SqlCommand(salesinsertsearchquery, connection);
                try
                {
                    ConnectionCheck();
                    com.ExecuteNonQuery();
                    MessageBox.Show("Entry successfully inserted. ", "Report", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    table_relations();
                    Clearcontrols();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\n\n" + ex, "Error");
                }
            }
        }

        private void productsaleclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void productmiscellaneousfinequantitytextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(productmiscellaneousfinequantitytextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                productmiscellaneousfinequantitytextBox.Text = numberOnly;
                productmiscellaneousfinequantitytextBox.SelectionStart = productmiscellaneousfinequantitytextBox.Text.Length;
            }
            else
            {
                productmiscellaneousfinequantitytextBox.Text = "0";
            }
        }
        private void productmiscellaneousdamagequantitytextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(productmiscellaneousdamagequantitytextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                productmiscellaneousdamagequantitytextBox.Text = numberOnly;
                productmiscellaneousdamagequantitytextBox.SelectionStart = productmiscellaneousdamagequantitytextBox.Text.Length;
            }
            else
            {
                productmiscellaneousdamagequantitytextBox.Text = "0";
            }
        }
        private void panel4_MouseHover(object sender, EventArgs e)
        {
            float num;
            if (!float.TryParse(productmiscellaneousfinequantitytextBox.Text, out num))
            { productmiscellaneousfinequantitytextBox.Text = "0"; }
            if (!float.TryParse(productmiscellaneousdamagequantitytextBox.Text, out num))
            { productmiscellaneousdamagequantitytextBox.Text = "0"; }
            productmiscellaneoustotalvaluetextBox.Text = (Convert.ToDouble(productmiscellaneousfinequantitytextBox.Text) * Convert.ToDouble(productmiscellaneousdamagequantitytextBox.Text)).ToString();
        }
        private void productmiscellaneousclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }
        private void productmiscellaneouslistcomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConnectionCheck();
            SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select TOP 1 unit_cost from product_received_table where product_id = (select product_id from product_table where upper (product_name) = upper('" + productmiscellaneouslistcomboBox.SelectedItem.ToString() + "')) ORDER BY date_of_receipt DESC;", connection);
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                productmiscellaneousunitpricetextBox.Text = rdr["unit_cost"].ToString();
            }
        }
        private void productmiscellaneousinsertbutton_Click(object sender, EventArgs e)
        {
            if (productmiscellaneoustotalvaluetextBox.Text == "0" || productmiscellaneouslistcomboBox.SelectedItem == null || String.IsNullOrEmpty(productmiscellaneousreasontextBox.Text) || (productmiscellaneousfinequantitytextBox.Text == "0" && productmiscellaneousdamagequantitytextBox.Text == "0"))
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else
            {
                ConnectionCheck();
                string miscellaneoustablequery = "SET DATEFORMAT dmy; Insert into product_miscellaneous_table (date, product_id, fine_quantity, damaged_quantity, reason, unit_cost) VALUES ('" + productmiscellaneousdateTimePicker.Value.ToShortDateString() + "',(Select product_id from product_table where upper(product_name) = upper('" + productmiscellaneouslistcomboBox.SelectedItem.ToString() + "'),'" + productmiscellaneousfinequantitytextBox.Text + "', '" + productmiscellaneousdamagequantitytextBox.Text + "', '" + productmiscellaneousreasontextBox.Text + "','" + productmiscellaneousunitpricetextBox.Text + "');";
                SqlCommand com = new SqlCommand(miscellaneoustablequery, connection);
                try
                {
                    ConnectionCheck();
                    com.ExecuteNonQuery();
                    MessageBox.Show("Product updated successfully.", "Success");
                    table_relations();
                    Clearcontrols();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update Product miscellaneous table! " + ex, "Error");
                }
            }
        }

        private void productdamagenamecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConnectionCheck();
            SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select TOP 1 unit_cost from product_received_table where product_id = (select product_id from product_table where upper (product_name) = upper('" + productdamagenamecomboBox.SelectedItem.ToString() + "')) ORDER BY date_of_receipt DESC;", connection);
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                productdamageunitpricetextBox.Text = rdr["unit_cost"].ToString();
            }
        }

        private void productdamagequantitytextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(productdamagequantitytextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                productdamagequantitytextBox.Text = numberOnly;
                productdamagequantitytextBox.SelectionStart = productdamagequantitytextBox.Text.Length;
            }
            else
            {
                productdamagequantitytextBox.Text = "";
            }
        }

        private void damageproductclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }
        private void damageproductinsertbutton_Click(object sender, EventArgs e)
        {
            if (productdamagetotalvaluetextBox.Text == "0" || productdamagenamecomboBox.SelectedItem == null || String.IsNullOrEmpty(productdamagequantitytextBox.Text))
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else
            {

                ConnectionCheck();
                string discardtablequery = "SET DATEFORMAT dmy; Insert into product_damage_table (date, product_id, period_of_damage, quantity_damaged, unit_cost) VALUES('" + newproductdateTimePicker.Value.ToShortDateString() + "',(Select product_id from product_table where upper(product_name) = upper('" + productdamagenamecomboBox.SelectedItem.ToString() + "'),'After Receipt','" + damagequantityreceipttextBox.Text + "', '" + productdamageunitpricetextBox.Text + "');";
                SqlCommand com = new SqlCommand(discardtablequery, connection);
                try
                {
                    ConnectionCheck();
                    com.ExecuteNonQuery();
                    MessageBox.Show("Product updated successfully.", "Success");
                    table_relations();
                    Clearcontrols();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update Product damage table! " + ex, "Error");
                }
            }
        }

        private void productsalecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConnectionCheck();
            SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select T0P 1 A.current_unit_price current_unit_price, B.unit_cost unit_cost from product_table as A, inner join product_received_table as B ON A.product_id = B.product_id where upper (A.product_name) = upper('" + productdamagenamecomboBox.SelectedItem.ToString() + "') ORDER BY B.date_of_receipt DESC;", connection);
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                currentunitpricetextBox.Text = rdr["current_unit_price"].ToString();
                currentunitcosttextBox.Text = rdr["unit_cost"].ToString();
            }
        }

        private void panel5_MouseHover(object sender, EventArgs e)
        {
            float num;
            if (!float.TryParse(productdamagequantitytextBox.Text, out num))
            { productdamagequantitytextBox.Text = "0"; }
            productdamagetotalvaluetextBox.Text = (Convert.ToDouble(productdamagequantitytextBox.Text) * Convert.ToDouble(productdamageunitpricetextBox.Text)).ToString();

        }

        private void creditpaymentamounttextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            string numberOnly = Regex.Replace(creditpaymentamounttextBox.Text, "[^0-9.+-]", "");
            creditpaymentamounttextBox.Text = numberOnly;
            creditpaymentamounttextBox.SelectionStart = creditpaymentamounttextBox.Text.Length;
        }

        private void creditpaymentamounttextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(creditpaymentamounttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                creditpaymentamounttextBox.Text = numberOnly;
                creditpaymentamounttextBox.SelectionStart = creditpaymentamounttextBox.Text.Length;
            }
            else
            {
                creditpaymentamounttextBox.Text = "";
            }
        }

        private void panel6_MouseHover(object sender, EventArgs e)
        {
            float num;
            if (!float.TryParse(creditpaymentamounttextBox.Text, out num) || (Convert.ToDouble(creditcurrentamounttextBox.Text) - Convert.ToDouble(creditpaymentamounttextBox.Text)) > 0)
            { creditpaymentamounttextBox.Text = "0"; }
            creditbalanceamounttextBox.Text = (Convert.ToDouble(creditcurrentamounttextBox.Text) - Convert.ToDouble(creditpaymentamounttextBox.Text)).ToString();
            ConnectionCheck();
            SqlCommand command2 = new SqlCommand("SET DATEFORMAT dmy; Select sum(total_credit_amount) total_credit_amount, sum(current_credit_amount) current_credit_amount from product_customer_credit_table;", connection);
            SqlDataReader reader2 = command2.ExecuteReader();
            if (!reader2.Read())
            {
                totalcreditamountlabellabel.Text = "0";
                currentcreditamountlabel.Text = "0";
            }
            else
            {
                totalcreditamountlabellabel.Text = reader2["total_credit_amount"].ToString();
                currentcreditamountlabel.Text = reader2["current_credit_amount"].ToString();
                if (String.IsNullOrEmpty(totalcreditamountlabellabel.Text))
                {
                    totalcreditamountlabellabel.Text = "0";
                }
                if (String.IsNullOrEmpty(currentcreditamountlabel.Text))
                {
                    currentcreditamountlabel.Text = "0";
                }
            }
        }

        private void creditcustomernamecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ConnectionCheck();
            SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select (total_credit_amount-amount_paid) current_credit_amount from product_customer_credit_table where customer_name = '" + creditcustomernamecomboBox.SelectedItem.ToString() + "';", connection);
            SqlDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                creditcurrentamounttextBox.Text = rdr["current_credit_amount"].ToString();
            }
        }

        private void creditdepositclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void creditdepositinsertbutton_Click(object sender, EventArgs e)
        {
            if (creditpaymentamounttextBox.Text == "0" || creditcustomernamecomboBox.SelectedItem == null || receiptidtextBox.ToString() == "" || creditpaymentamounttextBox.Text == "0")
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else
            {
                ConnectionCheck();
                string creditdepositquery = "SET DATEFORMAT dmy; Insert into product_credit_deposit_table (date_of_deposit, receipt_id, amount_paid, customer_name) VALUES('" + newproductdateTimePicker.Value.ToShortDateString() + "', '" + receiptidtextBox.Text + "','" + creditpaymentamounttextBox.Text + "','" + creditcustomernamecomboBox.SelectedItem.ToString() + "');";
                creditdepositquery += "SET DATEFORMAT dmy; update product_customer_credit_table set amount_paid = (amount_paid+" + Convert.ToInt32(creditpaymentamounttextBox.Text) + ") where upper(customer_name) = upper('" + creditcustomernamecomboBox.SelectedItem.ToString() + "'); ";
                SqlCommand com = new SqlCommand(creditdepositquery, connection);
                try
                {
                    ConnectionCheck();
                    com.ExecuteNonQuery();
                    MessageBox.Show("Credit value updated successfully.", "Success");
                    table_relations();
                    Clearcontrols();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update credit! " + ex, "Error");
                }
            }
        }
        private void table_relations()
        {
            ConnectionCheck();
            SqlCommand rmcmd8 = new SqlCommand("Select raw_material_name from raw_material_table", connection);
            ArrayList rmnamelist = new ArrayList();
            SqlDataReader rmnamerdr = rmcmd8.ExecuteReader();
            while (rmnamerdr.Read())
            {
                rmnamelist.Add(rmnamerdr["raw_material_name"].ToString());
            }
            int rmnamecount = rmnamelist.Count;
            for (int c = 0; c < rmnamecount; c++)
            {
                SqlCommand rmnamecmd3 = new SqlCommand("SET DATEFORMAT dmy; Update raw_material_current_cost_table set current_unit_cost = (select TOP 1 unit_cost from raw_material_purchase_table where rm_id = (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rmnamelist[c] + "')) ORDER BY date_of_purchase DESC) WHERE upper(raw_material_name) = upper('" + rmnamelist[c] + "') ", connection);
                ConnectionCheck();
                rmnamecmd3.ExecuteNonQuery();
            }
            ConnectionCheck();
            ArrayList productid_array = new ArrayList();
            //int[] productid_array;
            int productidsum = 0;
            SqlCommand command = new SqlCommand("select product_id from product_table;", connection);
            SqlDataReader productreader = command.ExecuteReader();
            while (productreader.Read())
            {
                productid_array.Add(productreader["product_id"].ToString());
            }
            productidsum = productid_array.Count;
            string fine_product_recorded = "";
            string damage_product_recorded = "";
            string fine_product_sold = "";
            string damaged_product_sold = "";
            string fine_miscellaneous_product = "";
            string damaged_miscellaneous_product = "";
            string damage_discarded = "0";
            string current_month_product_received = "";
            string current_month_product_sold = "";
            string average_unit_price = "";
            string average_unit_cost = "";
            if (productidsum != 0)
            {
                for (int a = 0; a < productidsum; a++)
                {
                    string product_tablequery1 = "SET DATEFORMAT dmy; Select ((sum(A.fine_quantity_received + A.damaged_quantity_received))-(sum(B.quantity_damaged))) fine_producted_recorded from [dbo].[product_received_table] AS A, [dbo].[product_damage_table] AS B where A.product_id = B.product_id and A.product_id = '" + productid_array[a] + "' group by A.product_id;";
                    string product_tablequery2 = "SET DATEFORMAT dmy; select sum(quantity_damaged) damage_product_recorded from [dbo].[product_damage_table] where product_id = '" + productid_array[a] + "' group by product_id;";
                    string product_tablequery3 = "SET DATEFORMAT dmy; Select sum(fine_quantity_sold) fine_product_sold, sum(damaged_quantity_sold) damaged_product_sold from [dbo].[product_sales_table] where product_id = '" + productid_array[a] + "' group by product_id;";
                    string product_tablequery4 = "SET DATEFORMAT dmy; Select sum(fine_quantity) fine_miscellaneous_product, sum(damaged_quantity) damaged_miscellaneous_product from [dbo].[product_miscellaneous_table] where product_id = '" + productid_array[a] + "' group by product_id;";
                    //string product_tablequery5 = "Select sum(damaged_discarded) damage_discarded from [dbo].[product_dicarded_table] where product_id = '" + productid_array[a] + "' andproduct_id = '" + productid_array[a] + "' group by product_id;";
                    string product_tablequery6 = "SET DATEFORMAT dmy; select sum(total_quantity_received) current_month_product_received from [dbo].[product_received_table] where upper(format(date_of_receipt, ('MMM-yyyy'))) = upper(format(getdate(), ('MMM-yyyy'))) and product_id = '" + productid_array[a] + "' group by format(date_of_receipt, ('MMM-yyyy')), product_id;";
                    string product_tablequery7 = "SET DATEFORMAT dmy; Select sum(total_quantity_sold) current_month_product_sold from [dbo].[product_sales_table] where upper(format(date_of_sale,('MMM-yyyy'))) = upper(format(getdate(),('MMM-yyyy'))) and product_id = '" + productid_array[a] + "'group by product_id;";
                    string product_tablequery8 = "SET DATEFORMAT dmy; Select ((sum(fine_unit_price + damaged_unit_price))/count(*)) average_unit_price from [dbo].[product_sales_table] where product_id = '" + productid_array[a] + "' group by product_id;";
                    string product_tablequery9 = "SET DATEFORMAT dmy; Select ((sum(unit_cost))/count(*)) average_unit_cost from [dbo].[product_received_table] where product_id = '" + productid_array[a] + "' group by product_id;";
                    try
                    {
                        ConnectionCheck();
                        SqlCommand producttablecom1 = new SqlCommand(product_tablequery1, connection);
                        SqlDataReader newreader1 = producttablecom1.ExecuteReader();
                        while (newreader1.Read())
                        {
                            fine_product_recorded = newreader1["fine_product_recorded"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom2 = new SqlCommand(product_tablequery2, connection);
                        SqlDataReader newreader2 = producttablecom2.ExecuteReader();
                        while (newreader2.Read())
                        {
                            damage_product_recorded = newreader2["damage_product_recorded"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom3 = new SqlCommand(product_tablequery3, connection);
                        SqlDataReader newreader3 = producttablecom3.ExecuteReader();
                        while (newreader3.Read())
                        {
                            fine_product_sold = newreader3["fine_product_sold"].ToString();
                            damaged_product_sold = newreader3["damaged_product_sold"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom4 = new SqlCommand(product_tablequery4, connection);
                        SqlDataReader newreader4 = producttablecom4.ExecuteReader();
                        while (newreader4.Read())
                        {
                            fine_miscellaneous_product = newreader4["fine_miscellaneous_product"].ToString();
                            damaged_miscellaneous_product = newreader4["damaged_miscellaneous_product"].ToString();
                        }
                        /*ConnectionCheck();
                        SqlCommand producttablecom5 = new SqlCommand(product_tablequery5, connection);
                        SqlDataReader newreader5 = producttablecom5.ExecuteReader();
                        while (newreader5.Read())
                        {
                            damage_discarded = newreader5["damage_discarded"].ToString();
                        }*/
                        ConnectionCheck();
                        SqlCommand producttablecom6 = new SqlCommand(product_tablequery6, connection);
                        SqlDataReader newreader6 = producttablecom6.ExecuteReader();
                        while (newreader6.Read())
                        {
                            current_month_product_received = newreader6["current_month_product_received"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom7 = new SqlCommand(product_tablequery7, connection);
                        SqlDataReader newreader7 = producttablecom7.ExecuteReader();
                        while (newreader7.Read())
                        {
                            current_month_product_sold = newreader7["current_month_product_sold"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom8 = new SqlCommand(product_tablequery8, connection);
                        SqlDataReader newreader8 = producttablecom8.ExecuteReader();
                        while (newreader8.Read())
                        {
                            average_unit_price = newreader8["average_unit_price"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand producttablecom9 = new SqlCommand(product_tablequery9, connection);
                        SqlDataReader newreader9 = producttablecom9.ExecuteReader();
                        while (newreader9.Read())
                        {
                            average_unit_cost = newreader9["average_unit_cost"].ToString();
                        }

                        string updatequery = "SET DATEFORMAT dmy; Update product_table set fine_product_recorded = '" + fine_product_recorded + "', damage_product_recorded = '" + damage_product_recorded + "', fine_product_sold = '" + fine_product_sold + "', damaged_product_sold = '" + damaged_product_sold + "', fine_miscellaneous_product = '" + fine_miscellaneous_product + "', damaged_miscellaneous_product = '" + damaged_miscellaneous_product + "', damage_discarded = '" + damage_discarded + "', current_month_product_received = '" + current_month_product_received + "', current_month_product_sold = '" + current_month_product_sold + "', average_unit_price = '" + average_unit_price + "', average_unit_cost = '" + average_unit_cost + "' where product_id = '" + productid_array[a] + "';";
                        SqlCommand com = new SqlCommand(updatequery, connection);
                        ConnectionCheck();
                        com.ExecuteNonQuery();
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show("Entry problem with table relations!!: " + exp, "Error", MessageBoxButtons.OK);
                    }

                }
                for (int a = 0; a < productidsum; a++)
                {
                    int month_checker1 = 0;
                    string month_of_receipt = "";
                    string quantity_received = "";
                    string total_cost = "";
                    string damaged_quantity_received = "";
                    string current_unit_price = "";
                    string monthlyreceivedtablequery1 = "SET DATEFORMAT dmy; select upper(format(A.date_of_receipt,('MMM-yyyy'))) month_of_receipt,  sum(A.total_quantity_received) quantity_received, sum(A.quantity_cost) total_cost, sum(B.quantity_damaged) damaged_quantity_received from  [dbo].[product_received_table] AS A, product_damage_table as B where product_id =  '" + productid_array[a] + "' and upper(format(A.date_of_receipt,('MMM-yyyy'))) = upper(format(B.date,('MMM-yyyy')))  group by format(A.date_of_receipt,('MMM-yyyy')), A.product_id";
                    string monthlyreceivedtablequery2 = "select  current_unit_price from[dbo].[product_table] where product_id =  '" + productid_array[a] + "';";
                    try
                    {
                        ConnectionCheck();
                        SqlCommand mrcom1 = new SqlCommand(monthlyreceivedtablequery2, connection);
                        SqlDataReader mrrdr1 = mrcom1.ExecuteReader();
                        while (mrrdr1.Read())
                        {
                            month_of_receipt = mrrdr1["month_of_receipt"].ToString();
                            quantity_received = mrrdr1["quantity_received"].ToString();
                            total_cost = mrrdr1["total_cost"].ToString();
                            damaged_quantity_received = mrrdr1["damaged_quantity_received"].ToString();

                        }
                        ConnectionCheck();
                        SqlCommand mrcom2 = new SqlCommand(monthlyreceivedtablequery2, connection);
                        SqlDataReader mrrdr2 = mrcom2.ExecuteReader();
                        while (mrrdr2.Read())
                        {
                            current_unit_price = mrrdr2["current_unit_price"].ToString();
                        }
                        string month_checker = "SET DATEFORMAT dmy; Select count(*) checka from monthly_product_received_table where upper(format(A.date_of_receipt,('MMM-yyyy'))) = upper('" + month_of_receipt + "');";
                        ConnectionCheck();
                        SqlCommand com6 = new SqlCommand(month_checker, connection);
                        string stringmreceivedupdatequery = "";
                        SqlDataReader rdr6 = com6.ExecuteReader();
                        while (rdr6.Read()) { month_checker1 = Convert.ToInt32(rdr6["checka"]); }
                        if (month_checker1 == 1)
                        { stringmreceivedupdatequery = "SET DATEFORMAT dmy; Update monthly_product_received_table set quantity_received = '" + quantity_received + "', quantity_cost = '" + total_cost + "', damaged_quantity_received = '" + damaged_quantity_received + "' where upper(month_of_receipt) = upper('" + month_of_receipt + "') and product_id = '" + productid_array[a] + "'; "; }
                        else if (month_checker1 == 0)
                        { stringmreceivedupdatequery = "SET DATEFORMAT dmy; Insert into monthly_product_received_table (month_of_receipt, product_id, quantity_received, quantity_cost, estimated_unit_price, damaged_quantity_received) VALUES (uppper('" + month_of_receipt + "'), '" + productid_array[a] + "', '" + quantity_received + "', '" + total_cost + "', '" + current_unit_price + "', '" + damaged_quantity_received + "');"; }
                        else
                        {
                            MessageBox.Show("Problem with first Monthly table!", "Error");
                        }
                        ConnectionCheck();
                        SqlCommand com7 = new SqlCommand(stringmreceivedupdatequery, connection);
                        com7.ExecuteNonQuery();


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Problem with first Monthly table! " + ex, "Error");
                    }
                }
                for (int a = 0; a < productidsum; a++)
                {
                    int month_checker1 = 0;
                    string month_of_sale = "";
                    string fine_quantity_sold = "";
                    string fine_quantity_total_income = "";
                    string damaged_quantity_sold = "";
                    string damaged_quantity_total_income = "";
                    string credit_sales = "";
                    string cash_sales = "";
                    string monthlysalestablequery1 = "select upper(format(date_of_sale, ('MMM-yyyy'))) month_of_sale, sum(fine_quantity_sold) fine_quantity_sold, sum(fine_total_amount) fine_quantity_total_income, sum(damaged_quantity_sold) damaged_quantity_sold, sum(damaged_total_amount) damaged_quantity_total_income from [dbo].[product_sales_table] where product_id = '" + productid_array[a] + "' group by format(date_of_sale,('MMM-yyyy')) order by format(date_of_sale,('MMM-yyyy')) desc;";
                    string monthlysalestablequery2 = "Select sum(total_amount) credit_sales from product_sales_table where upper(payment_type) = 'CREDIT'  group by format(date_of_sale,('MMM-yyyy')) and product_id = '" + productid_array[a] + "' order by format(date_of_sale,('MMM-yyyy')) desc";
                    string monthlysalestablequery3 = "Select sum(total_amount) cash_sales from product_sales_table where upper(payment_type) = 'CASH'  group by format(date_of_sale,('MMM-yyyy')) and product_id = '" + productid_array[a] + "' order by format(date_of_sale,('MMM-yyyy')) desc";
                    try
                    {
                        ConnectionCheck();
                        SqlCommand mscom1 = new SqlCommand(monthlysalestablequery1, connection);
                        SqlDataReader msrdr1 = mscom1.ExecuteReader();
                        while (msrdr1.Read())
                        {
                            month_of_sale = msrdr1["month_of_sale"].ToString();
                            fine_quantity_sold = msrdr1["fine_quantity_sold"].ToString();
                            fine_quantity_total_income = msrdr1["fine_quantity_total_income"].ToString();
                            damaged_quantity_sold = msrdr1["damaged_quantity_sold"].ToString();
                            damaged_quantity_total_income = msrdr1["damaged_quantity_total_income"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand mscom2 = new SqlCommand(monthlysalestablequery2, connection);
                        SqlDataReader msrdr2 = mscom2.ExecuteReader();
                        while (msrdr2.Read())
                        {
                            credit_sales = msrdr2["credit_sales"].ToString();
                        }
                        ConnectionCheck();
                        SqlCommand mscom3 = new SqlCommand(monthlysalestablequery3, connection);
                        SqlDataReader msrdr3 = mscom3.ExecuteReader();
                        while (msrdr3.Read())
                        {
                            cash_sales = msrdr3["cash_sales"].ToString();
                        }

                        string month_checker = "Select count(*) checka from monthly_product_sales_table where upper(format(A.date_of_sale,('MMM-yyyy'))) = upper('" + month_of_sale + "');";
                        ConnectionCheck();
                        SqlCommand com6 = new SqlCommand(month_checker, connection);
                        string stringmsalesupdatequery = "";
                        SqlDataReader rdr6 = com6.ExecuteReader();
                        while (rdr6.Read()) { month_checker1 = Convert.ToInt32(rdr6["checka"]); }
                        if (month_checker1 == 1)
                        { stringmsalesupdatequery = "SET DATEFORMAT dmy; Update monthly_sales_received_table set fine_quantity_sold = '" + fine_quantity_sold + "', fine_quantity_total_income = '" + fine_quantity_total_income + "', damaged_quantity_sold = '" + damaged_quantity_sold + "', damaged_quantity_total_income = '" + damaged_quantity_total_income + "', cash_sales = '" + cash_sales + "', credit_sales = '" + credit_sales + "' where month of receipt = upper('" + damaged_quantity_total_income + "') and product_id = '" + productid_array[a] + "'; "; }
                        else if (month_checker1 == 0)
                        { stringmsalesupdatequery = "SET DATEFORMAT dmy; Insert into monthly_product_sales_table (month_of_sale, product_id, fine_quantity_sold, fine_quantity_total_income, damaged_quantity_sold, damaged_total_income, cash_sales, credit_sales) VALUES (uppper('" + month_of_sale + "'), '" + productid_array[a] + "', '" + fine_quantity_sold + "', '" + fine_quantity_total_income + "', '" + damaged_quantity_sold + "', '" + damaged_quantity_total_income + "', '" + cash_sales + "', '" + credit_sales + "');"; }
                        else
                        {
                            MessageBox.Show("Problem with second Monthly table!", "Error");
                        }
                        ConnectionCheck();
                        SqlCommand com8 = new SqlCommand(stringmsalesupdatequery, connection);
                        com8.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Problem with seecond Monthly table! " + ex, "Error");
                    }
                }
            }
            ConnectionCheck();
            ArrayList raw_materialid_array = new ArrayList();
            int raw_materialidsum = 0;
            SqlCommand rmcommand = new SqlCommand("select rm_id from raw_material_table;", connection);
            SqlDataReader rmreader = rmcommand.ExecuteReader();
            while (rmreader.Read())
            {
                raw_materialid_array.Add(rmreader["rm_id"].ToString());
            }
            raw_materialidsum = raw_materialid_array.Count;
            string rmtableupdatequery = "";
            if (raw_materialidsum != 0)
            {
                for (int b = 0; b < raw_materialidsum; b++)
                {

                    rmtableupdatequery = "SET DATEFORMAT dmy; update raw_material_table set total_quantity_used = ISNULL((Select sum(quantity_used) from raw_material_to_product_usage_table where rm_id = '" + raw_materialid_array[b] + "' group by rm_id),0), total_quantity_bought = ISNULL((Select sum(quantity_purchased) from raw_material_purchase_table where rm_id = '" + raw_materialid_array[b] + "' group by rm_id ),0), total_quantity_discarded = ISNULL((Select sum(quantity_discarded) from raw_material_discarded_table where rm_id = '" + raw_materialid_array[b] + "' group by rm_id ),0) where rm_id = '" + raw_materialid_array[b] + "';";
                    SqlCommand rmcom2 = new SqlCommand(rmtableupdatequery, connection);
                    try
                    {
                        ConnectionCheck();
                        rmcom2.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Problem with RM table update! \n\n" + ex, "Error");
                    }
                }
            }
        }

        private void rawmaterialpurchaseunitcosttextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialpurchaseunitcosttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rawmaterialpurchaseunitcosttextBox.Text = numberOnly;
                rawmaterialpurchaseunitcosttextBox.SelectionStart = rawmaterialpurchaseunitcosttextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchaseunitcosttextBox.Text = "";
            }
        }

        private void rawmaterialpurchasequantityreceivedtextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialpurchasequantityreceivedtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = numberOnly;
                rawmaterialpurchasequantityreceivedtextBox.SelectionStart = rawmaterialpurchasequantityreceivedtextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = "";
            }
        }

        private void rawmaterialpurchaseeditunitcosttextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialpurchaseeditunitcosttextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rawmaterialpurchaseeditunitcosttextBox.Text = numberOnly;
                rawmaterialpurchaseeditunitcosttextBox.SelectionStart = rawmaterialpurchaseeditunitcosttextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
        }

        private void rawmaterialpurchasenamecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rawmaterialpurchasenamecomboBox.SelectedItem.ToString() == "Other")
            {
                otherrawmaterialpurchasenametextBox.ReadOnly = false;
                otherrawmaterialpurchasenametextBox.Text = "";
                rawmaterialpurchaseunitofmeasurementtextBox.ReadOnly = false;
                rawmaterialpurchaseunitofmeasurementtextBox.Text = "";
                rawmaterialpurchaseunitcosttextBox.ReadOnly = false;
                rawmaterialpurchaseunitcosttextBox.Text = "";
                editrawmaterialunitcostcheckBox.Enabled = false;
                editrawmaterialunitcostcheckBox.Checked = false;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
            }
            else if (rawmaterialpurchasenamecomboBox.SelectedItem.ToString() == "")
            {
                otherrawmaterialpurchasenametextBox.Text = "";
                otherrawmaterialpurchasenametextBox.ReadOnly = true;
                rawmaterialpurchaseunitofmeasurementtextBox.Text = "";
                rawmaterialpurchaseunitofmeasurementtextBox.ReadOnly = true;
                rawmaterialpurchaseunitcosttextBox.ReadOnly = true;
                rawmaterialpurchaseunitcosttextBox.Text = "";
                editrawmaterialunitcostcheckBox.Enabled = false;
                editrawmaterialunitcostcheckBox.Checked = false;
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
            else
            {
                ConnectionCheck();
                SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select TOP 1 A.unit_of_measurement,B.unit_cost from raw_material_table as A, raw_material_purchase_table as B where A.rm_id = B.rm_id and upper(A.raw_material_name) = upper('" + rawmaterialpurchasenamecomboBox.SelectedItem.ToString() + "') ORDER BY B.date_of_purchase DESC;", connection);
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    rawmaterialpurchaseunitofmeasurementtextBox.Text = rdr["unit_of_measurement"].ToString();
                    rawmaterialpurchaseunitcosttextBox.Text = rdr["unit_cost"].ToString();
                }
                editrawmaterialunitcostcheckBox.Enabled = true;
                editrawmaterialunitcostcheckBox.Checked = false;
            }

        }
        private void panel7_MouseHover(object sender, EventArgs e)
        {
            float num;
            if (float.TryParse(rawmaterialpurchaseunitcosttextBox.Text, out num))
            {
                rawmaterialpurchaseunitcosttextBox.Text = rawmaterialpurchaseunitcosttextBox.Text;
                rawmaterialpurchaseunitcosttextBox.SelectionStart = rawmaterialpurchaseunitcosttextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchaseunitcosttextBox.Text = "0";
            }
            if (float.TryParse(rawmaterialpurchasequantityreceivedtextBox.Text, out num))
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = rawmaterialpurchasequantityreceivedtextBox.Text;
                rawmaterialpurchasequantityreceivedtextBox.SelectionStart = rawmaterialpurchasequantityreceivedtextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchasequantityreceivedtextBox.Text = "0";
            }
            if (float.TryParse(rawmaterialpurchaseeditunitcosttextBox.Text, out num))
            {
                rawmaterialpurchaseeditunitcosttextBox.Text = rawmaterialpurchaseeditunitcosttextBox.Text;
                rawmaterialpurchaseeditunitcosttextBox.SelectionStart = rawmaterialpurchaseeditunitcosttextBox.Text.Length;
            }
            else
            {
                rawmaterialpurchaseeditunitcosttextBox.Text = "0";
            }

            if (editrawmaterialunitcostcheckBox.Enabled == true && editrawmaterialunitcostcheckBox.Checked == true && rawmaterialpurchaseunitcosttextBox.ReadOnly == true && String.IsNullOrEmpty(rawmaterialpurchaseunitcosttextBox.Text))
            {
                if (!float.TryParse(rawmaterialpurchaseeditunitcosttextBox.Text, out num)) { rawmaterialpurchaseeditunitcosttextBox.Text = "0"; }
                if (!float.TryParse(rawmaterialpurchasequantityreceivedtextBox.Text, out num)) { rawmaterialpurchasequantityreceivedtextBox.Text = "0"; }
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = false;
                rawmaterialpurchasetotalcosttextBox.Text = (Convert.ToDouble(rawmaterialpurchaseeditunitcosttextBox.Text) * Convert.ToDouble(rawmaterialpurchasequantityreceivedtextBox.Text)).ToString();
            }
            else
            {
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
                if (!float.TryParse(rawmaterialpurchaseunitcosttextBox.Text, out num)) { rawmaterialpurchaseunitcosttextBox.Text = "0"; }
                if (!float.TryParse(rawmaterialpurchasequantityreceivedtextBox.Text, out num)) { rawmaterialpurchasequantityreceivedtextBox.Text = "0"; }
                rawmaterialpurchasetotalcosttextBox.Text = (Convert.ToDouble(rawmaterialpurchaseunitcosttextBox.Text) * Convert.ToDouble(rawmaterialpurchasequantityreceivedtextBox.Text)).ToString();

            }

        }

        private void rawmaterialpurchaseclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }
        private void rawmaterialpurchaseinsertbutton_Click(object sender, EventArgs e)
        {
            string rawmaterialpurchasequery = "";
            if (rawmaterialpurchasetotalcosttextBox.Text == "0" || rawmaterialpurchasenamecomboBox.SelectedItem == null || rawmaterialpurchasenamecomboBox.SelectedIndex == 0 || String.IsNullOrEmpty(rawmaterialpurchaseunitofmeasurementtextBox.Text) || String.IsNullOrEmpty(rawmaterialpurchaseunitcosttextBox.Text) || String.IsNullOrEmpty(rawmaterialpurchasequantityreceivedtextBox.Text))
            {
                MessageBox.Show("Please fill the all the fields", "Notice");
            }
            else if (rawmaterialpurchasenamecomboBox.SelectedItem.ToString() == "Other" && String.IsNullOrEmpty(otherrawmaterialpurchasenametextBox.Text))
            {
                MessageBox.Show("Please fill the all the fields", "Notice");
            }
            else if (editrawmaterialunitcostcheckBox.Checked == true && String.IsNullOrEmpty(rawmaterialpurchaseeditunitcosttextBox.Text))
            {
                MessageBox.Show("Please fill the all the fields", "Notice");
            }
            else
            {
                if (rawmaterialpurchasenamecomboBox.SelectedItem.ToString() == "Other")
                {
                    ConnectionCheck();
                    int counta = 0;
                    SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select count(*) counta from raw_material_table where upper(raw_material_name) = upper('" + otherrawmaterialpurchasenametextBox.Text + "')", connection);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        counta = Convert.ToInt32(rdr["counta"]);
                    }
                    if (counta != 0)
                    {
                        MessageBox.Show("Raw Material already exists!", "Notice");
                        return;
                    }
                    else
                    {
                        rawmaterialpurchasequery = "SET DATEFORMAT dmy; Insert into raw_material_table (date_of_added, raw_material_name, raw_material_desc, unit_of_measurement) VALUES('" + rawmaterialpurchasedate.Value.ToShortDateString() + "', upper('" + otherrawmaterialpurchasenametextBox.Text + "'),'" + rawmaterialpurchasedesctextBox.Text + "','" + rawmaterialpurchaseunitofmeasurementtextBox.Text + "');";
                        rawmaterialpurchasequery += "Select MAX(rm_id) rm_id from raw_material_table;";
                    }
                }
                else
                {
                    rawmaterialpurchasequery = "Select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rawmaterialpurchasenamecomboBox.SelectedItem.ToString() + "');";
                }
                int rmid = 0;
                SqlCommand query_table = new SqlCommand(rawmaterialpurchasequery, connection);
                try
                {
                    ConnectionCheck();
                    rmid = (Int32)query_table.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update RM table! \n " + rawmaterialpurchasequery + "\n\n" + ex, "Error");
                    return;
                }
                string rawmaterialpurchasetablequery = "SET DATEFORMAT dmy; Insert into raw_material_purchase_table (date_of_purchase, rm_id, quantity_purchased, unit_cost) VALUES ('" + rawmaterialpurchasedate.Value.ToShortDateString() + "','" + rmid + "', '" + rawmaterialpurchasequantityreceivedtextBox.Text + "','" + rawmaterialpurchaseunitcosttextBox.Text + "');";
                if (rmid != 0)
                {
                    SqlCommand com = new SqlCommand(rawmaterialpurchasetablequery, connection);
                    try
                    {
                        ConnectionCheck();
                        com.ExecuteNonQuery();
                        string meee = "";
                        double costaa = 0;
                        if (String.IsNullOrEmpty(otherrawmaterialpurchasenametextBox.Text)) { meee = rawmaterialpurchasenamecomboBox.SelectedItem.ToString(); }
                        else { meee = otherrawmaterialpurchasenametextBox.Text; }
                        if(editrawmaterialunitcostcheckBox.Checked == true) { costaa = Convert.ToDouble(rawmaterialpurchaseeditunitcosttextBox.Text); }
                        else { costaa = Convert.ToDouble(rawmaterialpurchaseunitcosttextBox.Text); }
                        string rmcosttble = "SET DATEFORMAT dmy; IF EXISTS (select * from raw_material_current_cost_table where upper(raw_material_name) = upper('" + meee + "')) BEGIN Update raw_material_current_cost_table  set current_unit_cost = '" + costaa + "' where upper(raw_material_name) = upper('" + meee + "') END;";
                         rmcosttble += "SET DATEFORMAT dmy; IF NOT EXISTS (select * from raw_material_current_cost_table where upper(raw_material_name) = upper('" + meee + "')) BEGIN INSERT INTO raw_material_current_cost_table (raw_material_name, current_unit_cost) VALUES (upper('" + meee + "'), '" + costaa + "') END";
                        SqlCommand rmcmd3 = new SqlCommand(rmcosttble, connection);
                        ConnectionCheck();
                        rmcmd3.ExecuteNonQuery();
                        MessageBox.Show("Raw Material updated successfully.", "Success");
                        table_relations();
                        Clearcontrols();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failure to update RM purchase table! " + ex, "Error");
                    }
                }
            }
        }

        private void rawmaterialquantitydiscardedtextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialpurchaseunitcosttextBox.Text, "[^0-9.+-]", "");
            rawmaterialpurchaseunitcosttextBox.Text = numberOnly;
            rawmaterialpurchaseunitcosttextBox.SelectionStart = rawmaterialpurchaseunitcosttextBox.Text.Length;
        }

        private void rawmaterialquantitydiscardedtextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rawmaterialquantitydiscardedtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rawmaterialquantitydiscardedtextBox.Text = numberOnly;
                rawmaterialquantitydiscardedtextBox.SelectionStart = rawmaterialquantitydiscardedtextBox.Text.Length;
            }
            else
            {
                rawmaterialquantitydiscardedtextBox.Text = "";
            }
        }

        private void panel8_MouseHover(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(rawmaterialquantitydiscardedtextBox.Text)) { rawmaterialquantitydiscardedtextBox.Text = "0"; }
            if (String.IsNullOrEmpty(rawmaterialdiscardtotalexpensetextBox.Text)) { rawmaterialdiscardtotalexpensetextBox.Text = "0"; }
            rawmaterialdiscardtotalexpensetextBox.Text = (Convert.ToDouble(rawmaterialquantitydiscardedtextBox.Text) * Convert.ToDouble(rawmaterialunitcosttextBox.Text)).ToString();
        }

        private void rawmaterialnamecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rawmaterialnamecomboBox.SelectedIndex == 0 || rawmaterialnamecomboBox.SelectedIndex == -1 || rawmaterialnamecomboBox.SelectedIndex == 0) { }
            else
            {
                try
                {
                    ConnectionCheck();
                    SqlCommand cmd = new SqlCommand("Select current_unit_cost from raw_material_current_cost_table where upper(raw_material_name) = upper('" + rawmaterialnamecomboBox.SelectedItem.ToString() + "');", connection);
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        rawmaterialunitcosttextBox.Text = rdr["current_unit_cost"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error is: \n" + ex);
                }
            }
        }

        private void rawmaterialdiscardclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void rawmaterialdiscardinsertbutton_Click(object sender, EventArgs e)
        {
            if (rawmaterialnamecomboBox.SelectedItem == null || String.IsNullOrEmpty(rawmaterialquantitydiscardedtextBox.Text) || String.IsNullOrEmpty(rawmaterialdiscardreasontextBox.Text) || rawmaterialdiscardtotalexpensetextBox.Text == "0")
            {
                MessageBox.Show("Please enter all relevant fields", "Error");
            }
            else
            {
                ConnectionCheck();
                string rmdiscardtablequery = "SET DATEFORMAT dmy; Insert into raw_material_discarded_table (date, rm_id, quantity_discarded, unit_cost, reason) VALUES ('" + rawmaterialdiscarddateTime.Value.ToShortDateString() + "',(Select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rawmaterialnamecomboBox.SelectedItem.ToString() + "'),'" + rawmaterialquantitydiscardedtextBox.Text + "', '" + rawmaterialunitcosttextBox.Text + "', '" + rawmaterialdiscardreasontextBox + "');";
                SqlCommand com = new SqlCommand(rmdiscardtablequery, connection);
                try
                {
                    ConnectionCheck();
                    com.ExecuteNonQuery();
                    MessageBox.Show("Raw Material updated successfully.", "Success");
                    table_relations();
                    Clearcontrols();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failure to update RM discarded table! " + ex, "Error");
                }
            }
        }

        private void panel10_MouseHover(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(rmproductallocationunittextBox.Text)) { rmproductallocationunittextBox.Text = "0"; }
            if (String.IsNullOrEmpty(rmunitcosttextBox.Text)) { rmunitcosttextBox.Text = "0"; }
            rmproductallocationunitcosttextBox.Text = (Convert.ToDouble(rmproductallocationunittextBox.Text) * Convert.ToDouble(rmunitcosttextBox.Text)).ToString();
        }

        private void rmproductallocationproductnamecomboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rmproductallocationproductnamecomboBox.SelectedItem.ToString() == "Other")
            {
                newrmproductallocationproductnametextBox.ReadOnly = false;
                newrmproductallocationproductnametextBox.Text = "";
            }
            else
            {
                newrmproductallocationproductnametextBox.ReadOnly = true;
                newrmproductallocationproductnametextBox.Text = "";
            }
            rmunitcosttextBox.Text = "0";
            string rm_name = " ";
            if ((rmproductallocationrmnamecomboBox.SelectedItem == null)) { rm_name = " "; } else { rm_name = rmproductallocationrmnamecomboBox.SelectedItem.ToString(); }
            SqlCommand com6 = new SqlCommand("SET DATEFORMAT dmy; select TOP 1 unit_cost from raw_material_purchase_table where rm_id = (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "')) ORDER BY date_of_purchase DESC", connection);
            ConnectionCheck();
            SqlDataReader rdr6 = com6.ExecuteReader();
            while (rdr6.Read())
            {
                rmunitcosttextBox.Text = rdr6["unit_cost"].ToString();
            }
            if (String.IsNullOrEmpty(rmunitcosttextBox.Text)) { rmunitcosttextBox.Text = "0"; }
        }
        private void rmproductallocationclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void rmproductallocationunittextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(rmproductallocationunittextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                rmproductallocationunittextBox.Text = numberOnly;
                rmproductallocationunittextBox.SelectionStart = rmproductallocationunittextBox.Text.Length;
            }
            else
            {
                rmproductallocationunittextBox.Text = "0";
            }
        }

        private void rmproductallocationinsertbutton_Click(object sender, EventArgs e)
        {
            if (rmproductallocationrmnamecomboBox.SelectedItem == null || rmproductallocationunittextBox.Text == "0" || rmproductallocationunittextBox.Text == "" || rmproductallocationunitcosttextBox.Text == "0" && (rmproductallocationproductnamecomboBox.SelectedItem.ToString() == null || String.IsNullOrEmpty(newrmproductallocationproductnametextBox.Text)))
            {
                MessageBox.Show("Please input all fileds", "Notice");
            }
            else
            {
                string rmproductallocationclearbuttonquery = "";
                string productname = "";
                int checker = 0;
                if (rmproductallocationproductnamecomboBox.SelectedItem.ToString() == "Other")
                {
                    productname = newrmproductallocationproductnametextBox.Text;
                }
                else if ((rmproductallocationproductnamecomboBox.SelectedItem.ToString() == ""))
                {

                }
                else
                {
                    productname = rmproductallocationproductnamecomboBox.SelectedItem.ToString();
                }
                ConnectionCheck();
                SqlCommand cmd2 = new SqlCommand("Select count(*) counta from raw_material_product_allocation_table " +
                    "where upper(product_name) = upper('" + productname + "') and rm_id = (select rm_id from raw_material_table " +
                    "where upper(raw_material_name) = upper('" + rmproductallocationrmnamecomboBox.SelectedItem.ToString() + "'))", connection);
                SqlDataReader rdr2 = cmd2.ExecuteReader();
                while (rdr2.Read()) { checker = Convert.ToInt32(rdr2["counta"].ToString()); }
                if (checker != 0) { MessageBox.Show("Raw Material to product combinatiion exists!", "Error"); }
                else
                {
                    rmproductallocationclearbuttonquery = "SET DATEFORMAT dmy; " +
                        "Insert into raw_material_product_allocation_table (product_name, rm_id, rm_unit_quantity_utilized) " +
                        "VALUES (upper('" + productname + "'), (select rm_id from raw_material_table where upper(raw_material_name) " +
                        "= upper('" + rmproductallocationrmnamecomboBox.SelectedItem.ToString() + "')), '" + rmproductallocationunittextBox.Text + "');";
                    int tempchecker = 0;
                    SqlCommand cmd3 = new SqlCommand("Select count(product_name) counta " +
                        "from raw_m_product_temp_table where upper(product_name) = upper('" + productname + "');", connection);
                    ConnectionCheck();
                    SqlDataReader rdr3 = cmd3.ExecuteReader();
                    while (rdr3.Read())
                    {
                        tempchecker = Convert.ToInt32(rdr3["counta"].ToString());
                    }
                    if (tempchecker == 0)
                    {
                        rmproductallocationclearbuttonquery += "Insert into raw_m_temp_table (product_name) values (upper('" + productname + "'));";
                    }
                    try
                    {
                        ConnectionCheck();
                        SqlCommand cmd4 = new SqlCommand(rmproductallocationclearbuttonquery, connection);
                        cmd4.ExecuteNonQuery();
                        MessageBox.Show("Fields successfully inserted!", "Success");
                        table_relations();
                        Clearcontrols();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Fields NOT inserted in RMPA: " + ex, "Error");

                    }


                }
            }
        }

        private void rmproductallocationreportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("select A.raw_material_name raw_material_name, B.product_name product_name, " +
                                    "B.rm_unit_quantity_utilized rm_unit_quantity_utilized, B.rm_unit_quantity_utilized *" +
                                    " (select TOP 1  quantity_cost from Raw_material_purchase_table  where rm_id = A.rm_id Order by date_of_purchase desc) " +
                                    "unit_cost from raw_material_table as A, Raw_matrial_product_allocation_table as B where A.rm_id = B.rm_id", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Raw Material Name";
            sheet.Range["12"].Text = "Product Name";
            sheet.Range["C1"].Text = "Quantity utilized";
            sheet.Range["D1"].Text = "Quantitiy utilized cost";
            sheet.Range["A1:D1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["raw_material_name"].ToString();
                sheet.Range["B" + a].Text = rdr["product_name"].ToString();
                sheet.Range["C" + a].Text = rdr["rm_unit_quantity_utilized"].ToString();
                sheet.Range["D" + a].Text = rdr["unit_cost"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Raw Material Product Allocation Report.xlsx");

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Raw Material Product Allocation Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Raw Material Product Allocation Report.xlsx");
        }

        private void productmiscellaneousreportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("select B.date date, A.product_name product_name, A.unit_of_measurement unit_of_measurement, " +
                " (B.fine_quantity+B.damaged_quantity) total_quantity, B.unit_cost unit_cost, B.quantity_cost quantity_cost, B.reason reason " +
                "from product_table as A, product_miscellaneous_table as B where A.product_id = B.product_id order by B.date desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Date Recorded";
            sheet.Range["B1"].Text = "Product Name";
            sheet.Range["C1"].Text = "Unit of Measurement";
            sheet.Range["D1"].Text = "Quantitiy Recorded";
            sheet.Range["E1"].Text = "Unit Cost";
            sheet.Range["F1"].Text = "Total Cost";
            sheet.Range["G1"].Text = "Reason";
            sheet.Range["A1:G1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["date"].ToString();
                sheet.Range["B" + a].Text = rdr["product_name"].ToString();
                sheet.Range["C" + a].Text = rdr["unit_of_measurement"].ToString();
                sheet.Range["D" + a].Text = rdr["total_quantity"].ToString();
                sheet.Range["E" + a].Text = rdr["unit_cost"].ToString();
                sheet.Range["F" + a].Text = rdr["quantity_cost"].ToString();
                sheet.Range["G" + a].Text = rdr["reason"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Product Miscellaneous Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Product Miscellaneous Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Product Miscellaneous Report.xlsx");
        }

        private void rawmaterialdiscardbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("select B.date date, A.raw_material_name raw_material_name, A.unit_of_measurement unit_of_measurement,  " +
                "B.quantity_discarded quantity_discarded, B.unit_cost unit_cost, B.quantity_cost quantity_cost " +
                "from raw_material_table as A, raw_material_discarded_table as B where A.rm_id = B.rm_id order by B.date desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Date Discarded";
            sheet.Range["B1"].Text = "Raw Material Name";
            sheet.Range["C1"].Text = "Unit of Measurement";
            sheet.Range["D1"].Text = "Quantitiy discarded";
            sheet.Range["E1"].Text = "Unit Cost";
            sheet.Range["F1"].Text = "Total Cost";
            sheet.Range["A1:F1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["date"].ToString();
                sheet.Range["B" + a].Text = rdr["raw_material_name"].ToString();
                sheet.Range["C" + a].Text = rdr["unit_of_measurement"].ToString();
                sheet.Range["D" + a].Text = rdr["quantity_discarded"].ToString();
                sheet.Range["E" + a].Text = rdr["unit_cost"].ToString();
                sheet.Range["F" + a].Text = rdr["quantity_cost"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Raw Material Discard Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Raw Material Discard Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Raw Material Discard Report.xlsx");
        }

        private void productdamagereportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("select B.date date, A.product_name product_name, A.unit_of_measurement unit_of_measurement, " +
                " B.period_of_damage, B.quantity_damaged quantity_damaged from product_table as A, product_damage_table as B " +
                "where A.product_id = B.product_id order by B.date desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Date Recorded";
            sheet.Range["B1"].Text = "Product Name";
            sheet.Range["C1"].Text = "Unit of Measurement";
            sheet.Range["D1"].Text = "Quantitiy Recorded";
            sheet.Range["E1"].Text = "Unit Cost";
            sheet.Range["F1"].Text = "Total Cost";
            sheet.Range["A1:G1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["date"].ToString();
                sheet.Range["B" + a].Text = rdr["product_name"].ToString();
                sheet.Range["C" + a].Text = rdr["unit_of_measurement"].ToString();
                sheet.Range["D" + a].Text = rdr["total_quantity"].ToString();
                sheet.Range["E" + a].Text = rdr["unit_cost"].ToString();
                sheet.Range["F" + a].Text = rdr["quantity_cost"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Product Damage Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Product Damage Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Product Damage Report.xlsx");
        }

        private void previousbalancebutton_Click(object sender, EventArgs e)
        {
            DialogResult response = MessageBox.Show("Are you sure you want to record previous balance?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (response == DialogResult.Yes)
            {
                SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select product_id, fine_quantity_available, damage_quantity_available, " +
                    "current_unit_price, current_unit_cost from product_table;", connection);
                ConnectionCheck();
                SqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    string previousbal = "SET DATEFORMAT dmy; IF NOT EXISTS " +
                        "(Select * from previous_balance_table where upper(month_of_carry_over) = upper(format(dateadd(month , -1, getdate()),'MMM-yyyy'))) " +
                        "BEGIN Insert into Previous_balance_table (month_of_carry_over, product_id, fine_quantity_carried_over, " +
                        "damaged_quantity_carried_over, unit_price) VALUES  (upper(format(dateadd(month , -1, getdate()),'MMM-yyyy')), '" 
                        + rdr["product_id"].ToString() + "', '" + rdr["fine_quantity_available"].ToString() + "', '" 
                        + rdr["damage_quantity_available"].ToString() + "', '" + rdr["current_unit_price"].ToString() + "', '" 
                        + rdr["current_unit_cost"].ToString() + ")";
                    try
                    {
                        ConnectionCheck();
                        SqlCommand cmd2 = new SqlCommand(previousbal, connection);
                        cmd2.ExecuteNonQuery();
                        MessageBox.Show("Previous Balance Recorded", "", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Problem with the record!. " + ex, "Error", MessageBoxButtons.OK);
                    }
                }
            }
            else
            { MessageBox.Show("Operation Canceled!!", "Cancel", MessageBoxButtons.OK, MessageBoxIcon.Information); }

        }

        private void clientcreditreportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand totalcredittablequery = new SqlCommand("SET DATEFORMAT dmy; " +
                "select date_of_sale, voucher_id, upper(customer_name) customer_name, " +
                " credit_amount from product_credit_table order by date_of_sale desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = totalcredittablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet totalcredit = workbook.Worksheets[0];
            totalcredit.Name = "Total Credit";
            totalcredit.Range["A1"].Text = "Date Of Sale";
            totalcredit.Range["B1"].Text = "Voucher Id";
            totalcredit.Range["C1"].Text = "Customer Name";
            totalcredit.Range["D1"].Text = "Credit Amount";
            totalcredit.Range["A1:D1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                totalcredit.Range["A" + a].Text = rdr["date_of_sale"].ToString();
                totalcredit.Range["B" + a].Text = rdr["voucher_id"].ToString();
                totalcredit.Range["C" + a].Text = rdr["customer_name"].ToString();
                totalcredit.Range["D" + a].Text = rdr["credit_amount"].ToString();
                a++;
            }
            totalcredit.AllocatedRange.AutoFitColumns();
            totalcredit.AllocatedRange.AutoFitRows();

            SqlCommand currentcredittablequery = new SqlCommand("SET DATEFORMAT dmy; " +
                "select customer_name, total_credit_amount, amount_paid, current_credit_amount " +
                "from product_customer_credit_table where current_credit_amount != '0';", connection);
            ConnectionCheck();
            SqlDataReader rdr2 = currentcredittablequery.ExecuteReader();

            Worksheet currentcredit = workbook.Worksheets[1];
            currentcredit.Name = "Current Credit";
            currentcredit.Range["A1"].Text = "Customer Name";
            currentcredit.Range["B1"].Text = "Total Credit Amount";
            currentcredit.Range["C1"].Text = "Amount Paid";
            currentcredit.Range["D1"].Text = "Current Credit Amount";
            currentcredit.Range["A1:D1"].Style.Font.IsBold = true;
            int b = 2;
            while (rdr2.Read())
            {
                currentcredit.Range["A" + b].Text = rdr2["customer_name"].ToString();
                currentcredit.Range["B" + b].Text = rdr2["total_credit_amount"].ToString();
                currentcredit.Range["C" + b].Text = rdr2["amount_paid"].ToString();
                currentcredit.Range["D" + b].Text = rdr2["current_credit_amount"].ToString();
                b++;
            }
            currentcredit.AllocatedRange.AutoFitColumns();
            currentcredit.AllocatedRange.AutoFitRows();

            workbook.Worksheets.Remove("Sheet3");
            workbook.Worksheets[1].Activate();
            workbook.SaveToFile("Customer Credit Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Customer Credit Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[3]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Customer Credit Report.xlsx");
        }

        private void expensereportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("SET DATEFORMAT dmy; select date, expense_name, expense_type,  expense_detail, amount " +
                "from expenses_table order by date desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Date Recorded";
            sheet.Range["B1"].Text = "Expense Name";
            sheet.Range["C1"].Text = "Expense Type";
            sheet.Range["D1"].Text = "Expense Detail";
            sheet.Range["E1"].Text = "Amount";
            sheet.Range["A1:E1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["date"].ToString();
                sheet.Range["B" + a].Text = rdr["expense_name"].ToString();
                sheet.Range["C" + a].Text = rdr["expense_type"].ToString();
                sheet.Range["D" + a].Text = rdr["expense_detail"].ToString();
                sheet.Range["E" + a].Text = rdr["amount"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Expenses Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Expenses Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Expenses Report.xlsx");
        }
        private void previousbalancereportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("SET DATEFORMAT dmy; select B.month_of_carry_over, A.product_name, A.unit_of_measurement, " +
                "B.total_quantity_carried_over,  B.fine_quantity_carried_over, B.damaged_quantity_carried_over, B.unit_cost, " +
                "B.total_expense_carried_over from previous_balance_table as B, product_table as A " +
                "where A.product_id = B.product_id order by B.mpb_id desc", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Text = "Month";
            sheet.Range["B1"].Text = "Product Name";
            sheet.Range["C1"].Text = "Unit of Measurement";
            sheet.Range["D1"].Text = "Total Quantity";
            sheet.Range["E1"].Text = "Fine Qunatity";
            sheet.Range["F1"].Text = "Damaged Quantity";
            sheet.Range["G1"].Text = "Unit Cost";
            sheet.Range["H1"].Text = "Total Cost";
            sheet.Range["A1:H1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Text = rdr["month_of_carry_over"].ToString();
                sheet.Range["B" + a].Text = rdr["product_name"].ToString();
                sheet.Range["C" + a].Text = rdr["unit_of_measurement"].ToString();
                sheet.Range["D" + a].Text = rdr["total_quantity_carried_over"].ToString();
                sheet.Range["E" + a].Text = rdr["fine_quantity_carried_over"].ToString();
                sheet.Range["F" + a].Text = rdr["damaged_quantity_carried_over"].ToString();
                sheet.Range["G" + a].Text = rdr["unit_cost"].ToString();
                sheet.Range["H" + a].Text = rdr["total_expense_carried_over"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Previous Balance Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Previous Balance Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Previous Balance Report.xlsx");
        }

        private void currentmonthplreportbutton_Click(object sender, EventArgs e)
        {
            double num;
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.RightHeader = "Sample";
            sheet.PageSetup.LeftHeader = DateTime.Now.ToString("dd-MM-yy HH:mm:ss");
            CellRange range = sheet.Range["A1:D51"];
            range.BorderAround(LineStyleType.Thin, Color.Blue);
            //sheet.PageSetup.DifferentFirst = 1;

            sheet.Name = "Result";
            sheet.Range["A3"].Text = "REVENUE";
            sheet.Range["A17"].Text = "OVERHEAD";

            sheet.Range["B4"].Text = "Sales";
            sheet.Range["B9"].Text = "Total Sales";
            sheet.Range["B11"].Text = "Operating Expenses";
            sheet.Range["B13"].Text = "Total Cost";
            sheet.Range["B18"].Text = "Expenses";
            sheet.Range["B25"].Text = "Total Expenses";
            sheet.Range["B33"].Text = "2%";
            sheet.Range["B36"].Text = "50%";
            sheet.Range["B37"].Text = "25";
            sheet.Range["B38"].Text = "25%";


            sheet.Range["C5"].Text = "Cash Sales";
            sheet.Range["C6"].Text = "Credit Sales";
            sheet.Range["C7"].Text = "Deposits";
            sheet.Range["C8"].Text = "Damage";
            sheet.Range["C9"].Text = "Total Sales";
            sheet.Range["C12"].Text = "Cost of Goods Sold";
            sheet.Range["C15"].Text = "Gross Profit/Loss";
            sheet.Range["C19"].Text = "Salaries";
            sheet.Range["C20"].Text = "Electricity";
            sheet.Range["C21"].Text = "Assets";
            sheet.Range["C22"].Text = "Repairs";
            sheet.Range["C23"].Text = "Other Expenses";
            sheet.Range["C24"].Text = "Direct Expenses";
            sheet.Range["C27"].Text = "Net Profit/Loss";
            sheet.Range["C33"].Text = "2% of Stock";
            sheet.Range["C35"].Text = "Less 2% of Stock";
            sheet.Range["C36"].Text = "Person 1";
            sheet.Range["C37"].Text = "Person 2";
            sheet.Range["C38"].Text = "Person 3";
            sheet.Range["C42"].Text = "Closing Stock";
            sheet.Range["C44"].Text = "Mould & Machine";
            sheet.Range["C46"].Text = "Purchasing";

            double cash_sales = 0; // see below 
            double credit_sales = 0; // 
            SqlCommand com1 = new SqlCommand("SET DATEFORMAT dmy; " +
                "select cash_sales, credit_sales from monthly_product_sales_table " +
                "where month_of_sale  = upper(format(getdate(), ('MMM-yyyy')));", connection);
            string cash_sales1 = "";
            string credit_sales1 = "";
            ConnectionCheck();
            SqlDataReader rdr1 = com1.ExecuteReader();
            while (rdr1.Read())
            {
                cash_sales1 = rdr1["cash_sales"].ToString();
                credit_sales1 = rdr1["credit_sales"].ToString();

            }

            if (Double.TryParse(cash_sales1, out num))
            { cash_sales = Convert.ToDouble(cash_sales1); }
            if (Double.TryParse(credit_sales1, out num))
            { credit_sales = Convert.ToDouble(credit_sales1); }

            double deposits = 0; //
            SqlCommand com2 = new SqlCommand("SET DATEFORMAT dmy; select sum (amount_paid) deposits " +
                "from product_credit_deposit_table where upper(format(date_of_deposit, 'MMM-yyyy')) =  upper(format(getdate(), 'MMM-yyyy')) ;", connection);
            string deposits1 = "";
            ConnectionCheck();
            SqlDataReader rdr2 = com2.ExecuteReader();
            while (rdr2.Read())
            {
                deposits1 = rdr2["deposits"].ToString();
            }
            if (Double.TryParse(deposits1, out num))
            { deposits = Convert.ToDouble(deposits1); }

            double damage = 0;
            SqlCommand com3 = new SqlCommand("SET DATEFORMAT dmy; select sum(total_cost) damage " +
                "from product_damage_table where upper(format(date, 'MMM-yyyy')) =  upper(format(getdate(), 'MMM-yyyy'));", connection);
            string damage1 = "";
            ConnectionCheck();
            SqlDataReader rdr3 = com3.ExecuteReader();
            while (rdr3.Read())
            {
                damage1 = rdr3["damage"].ToString();

            }
            if (Double.TryParse(damage1, out num))
            { damage = Convert.ToDouble(damage1); }

            double total_sales = cash_sales + credit_sales + deposits + damage;

            double cost_of_goods_sold = 0;
            SqlCommand com4 = new SqlCommand("SET DATEFORMAT dmy; select sum(total_cost) cost_of_goods_sold " +
                "from product_sales_table where upper(format(date_of_sale, 'MMM-yyyy')) =  upper(format(getdate(), 'MMM-yyyy'));", connection);
            string cost_of_goods_sold1 = "";
            ConnectionCheck();
            SqlDataReader rdr4 = com4.ExecuteReader();
            while (rdr4.Read())
            {
                cost_of_goods_sold1 = rdr4["cost_of_goods_sold"].ToString();

            }
            if (Double.TryParse(cost_of_goods_sold1, out num))
            { cost_of_goods_sold = Convert.ToDouble(cost_of_goods_sold1); }

            double total_cost = cost_of_goods_sold;
            double gross_profit_loss = total_sales - total_cost;
            double salaries = 0;
            double electricity = 0;
            double assets = 0;
            double repairs = 0;
            double other_expenses = 0;
            double direct_expenses = 0;


            SqlCommand com5 = new SqlCommand("SET DATEFORMAT dmy; select sum(amount) electricity " +
                "from expenses_table where upper(format(date, 'MMM-yyyy')) = upper(format(getdate(), 'MMM-yyyy')) " +
                "and upper(expense_type) = 'ELECTRICITY';", connection);
            SqlCommand com6 = new SqlCommand("SET DATEFORMAT dmy; select sum(amount) salaries " +
                "from expenses_table where upper(format(date, 'MMM-yyyy')) =  upper(format(getdate(), 'MMM-yyyy')) " +
                "and upper(expense_type) = 'SALARY';", connection);
            string electricity1 = "";
            string salaries1 = "";
            ConnectionCheck();
            SqlDataReader rdr5 = com5.ExecuteReader();
            while (rdr5.Read())
            {
                electricity1 = rdr5["electricity"].ToString();
            }
            ConnectionCheck();
            SqlDataReader rdr6 = com6.ExecuteReader();
            while (rdr6.Read())
            {
                salaries1 = rdr6["salaries"].ToString();

            }

            if (Double.TryParse(electricity1, out num))
            { electricity = Convert.ToDouble(electricity1); }
            if (Double.TryParse(salaries1, out num))
            { salaries = Convert.ToDouble(salaries1); }

            double total_expenses = salaries + electricity + assets + repairs + other_expenses + direct_expenses;

            double net_profit_loss = gross_profit_loss - total_expenses;
            double avaialble_stock_cost = 0;

            SqlCommand com7 = new SqlCommand("Select sum(total_quantity_available*current_unit_cost) avaialble_stock_cost " +
                "from product_table", connection);
            string avaialble_stock_cost1 = "";
            ConnectionCheck();
            SqlDataReader rdr7 = com7.ExecuteReader();
            while (rdr7.Read())
            {
                avaialble_stock_cost1 = rdr7["avaialble_stock_cost"].ToString();
            }
            if (Double.TryParse(avaialble_stock_cost1, out num))
            { avaialble_stock_cost = Convert.ToDouble(avaialble_stock_cost1); }

            double two_percent_of_stock = avaialble_stock_cost * 0.02;
            double less_2_percent_stock = net_profit_loss - two_percent_of_stock;
            double person_1 = 0.5 * less_2_percent_stock;
            double person_2 = 0.2 * less_2_percent_stock;
            double person_3 = 0.2 * less_2_percent_stock;
            double closing_stock = avaialble_stock_cost;
            double mould_and_machine = 0;
            double purchasing = 0;

            SqlCommand com8 = new SqlCommand("Select sum(quantity_cost) purchasing from raw_material_purchase_table where upper(format(date_of_purchase, 'MMM-yyyy')) =  upper(format(getdate(), 'MMM-yyyy'))", connection);
            string purchasing1 = "";
            ConnectionCheck();
            SqlDataReader rdr8 = com8.ExecuteReader();
            while (rdr8.Read())
            {
                purchasing1 = rdr8["purchasing"].ToString();

            }
            if (Double.TryParse(purchasing1, out num))
            { purchasing = Convert.ToDouble(purchasing1); }

            sheet.Range["D5"].Text = "Ugx    " + cash_sales;
            sheet.Range["D6"].Text = "Ugx    " + credit_sales;
            sheet.Range["D7"].Text = "Ugx    " + deposits;
            sheet.Range["D8"].Text = "Ugx    " + damage;
            sheet.Range["D9"].Text = "Ugx    " + total_sales;
            sheet.Range["D12"].Text = "Ugx    " + cost_of_goods_sold;
            sheet.Range["D13"].Text = "Ugx    " + total_cost;
            sheet.Range["D15"].Text = "Ugx    " + gross_profit_loss;
            sheet.Range["D19"].Text = "Ugx    " + salaries;
            sheet.Range["D20"].Text = "Ugx    " + electricity;
            sheet.Range["D21"].Text = "Ugx    " + assets;
            sheet.Range["D22"].Text = "Ugx    " + repairs;
            sheet.Range["D23"].Text = "Ugx    " + other_expenses;
            sheet.Range["D24"].Text = "Ugx    " + direct_expenses;
            sheet.Range["D25"].Text = "Ugx    " + total_expenses;
            sheet.Range["D27"].Text = "Ugx    " + net_profit_loss;
            sheet.Range["D33"].Text = "Ugx    " + two_percent_of_stock;
            sheet.Range["D35"].Text = "Ugx    " + less_2_percent_stock;
            sheet.Range["D36"].Text = "Ugx    " + person_1;
            sheet.Range["D37"].Text = "Ugx    " + person_2;
            sheet.Range["D38"].Text = "Ugx    " + person_3;
            sheet.Range["D42"].Text = "Ugx    " + closing_stock;
            sheet.Range["D44"].Text = "Ugx    " + mould_and_machine;
            sheet.Range["D46"].Text = "Ugx    " + purchasing;

            sheet.Range["A3:B32"].Style.Font.IsBold = true;
            sheet.Range["C15"].Style.Font.IsBold = true;
            sheet.Range["C27"].Style.Font.IsBold = true;
            sheet.Range["C35"].Style.Font.IsBold = true;
            sheet.Range["C42"].Style.Font.IsBold = true;
            sheet.Range["C44"].Style.Font.IsBold = true;
            sheet.Range["C46"].Style.Font.IsBold = true;

            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Profit And Loss Report.xlsx", ExcelVersion.Version2013);

            workbook1 = xlApp.Workbooks.Open(programpath + @"\Profit And Loss Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);

            System.Diagnostics.Process.Start("Profit And Loss Report.xlsx");
        }

        private void mouldquantitytextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(mouldquantitytextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                mouldquantitytextBox.Text = numberOnly;
                mouldquantitytextBox.SelectionStart = mouldquantitytextBox.Text.Length;
            }
            else
            {
                mouldquantitytextBox.Text = "0";
            }
            mouldratetotaltextBox.Text = (Convert.ToDouble(mouldquantitytextBox.Text) * Convert.ToDouble(mouldratetextBox.Text)).ToString();
        }

        private void mouldratetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(mouldratetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                mouldratetextBox.Text = numberOnly;
                mouldratetextBox.SelectionStart = mouldratetextBox.Text.Length;
            }
            else
            {
                mouldratetextBox.Text = "0";
            }
            mouldratetotaltextBox.Text = (Convert.ToDouble(mouldquantitytextBox.Text) * Convert.ToDouble(mouldratetextBox.Text)).ToString();
        }

        private void mouldpricetextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(mouldpricetextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                mouldpricetextBox.Text = numberOnly;
                mouldpricetextBox.SelectionStart = mouldpricetextBox.Text.Length;
            }
            else
            {
                mouldpricetextBox.Text = "0";
            }
            mouldreceivedtotaltextBox.Text = (Convert.ToDouble(mouldpricetextBox.Text) * Convert.ToDouble(mouldreceivedtextBox.Text)).ToString();
        }

        private void mouldreceivedtextBox_TextChanged(object sender, EventArgs e)
        {
            string numberOnly = Regex.Replace(mouldreceivedtextBox.Text, "[^0-9.+-]", "");
            float num;
            if (float.TryParse(numberOnly, out num))
            {
                mouldreceivedtextBox.Text = numberOnly;
                mouldreceivedtextBox.SelectionStart = mouldreceivedtextBox.Text.Length;
            }
            else
            {
                mouldreceivedtextBox.Text = "0";
            }
            mouldreceivedtotaltextBox.Text = (Convert.ToDouble(mouldpricetextBox.Text) * Convert.ToDouble(mouldreceivedtextBox.Text)).ToString();
        }

        private void mouldclearbutton_Click(object sender, EventArgs e)
        {
            Clearcontrols();
        }

        private void mouldinsertbutton_Click(object sender, EventArgs e)
        {
            if (mouldnametextBox.Text.Length < 2 || moulddepartmenttextBox.Text.Length < 2 || mouldproductnametextBox.Text.Length < 2 || mouldratetextBox.Text == "0" || mouldratetextBox.Text == "" || mouldpricetextBox.Text == "0" || mouldpricetextBox.Text == "")
            {
                MessageBox.Show("Please fill all fields!!", "Notice", MessageBoxButtons.OK);
            }
            else
            {
                try
                {
                    ConnectionCheck();
                    /*SqlCommand cmd2 = new SqlCommand("Select (*) counta from mould_table where upper(mould_name) = upper('" + mouldnametextBox.Text + "');");
                    SqlDataReader rdr2 = cmd2.ExecuteReader();
                    int counta = 0;
                    while(rdr2.Read())
                    { counta = Convert.ToInt32(rdr2["counta"].ToString()); }
                    if (counta == 0)
                    {*/
                    SqlCommand cmd1 = new SqlCommand("SET DATEFORMAT dmy; Insert into mould_table (date_received, mould_name, product_name, department, mould_quantity, mould_rate, mould_price, mould_received) VALUES ('" + moulddateTimePicker.Value.ToShortDateString() + "', upper('" + mouldnametextBox.Text + "'), upper('" + mouldproductnametextBox.Text + "'), upper('" + moulddepartmenttextBox.Text + "'), '" + mouldquantitytextBox.Text + "', '" + mouldratetextBox.Text + "', '" + mouldpricetextBox.Text + "', '" + mouldreceivedtextBox.Text + "');", connection);
                    cmd1.ExecuteNonQuery();
                    MessageBox.Show("Record inserted successfully.", "Success", MessageBoxButtons.OK);
                    Clearcontrols();
                    this.mould_tableTableAdapter.Fill(this.mould_dbDataSet.Mould_table);
                    //}
                    /*else
                    {
                        SqlCommand cmd1 = new SqlCommand("Update mould_table set date_received = '" + moulddateTimePicker.Value.ToShortDateString() + "', mould_quantity = mould_quantity +'" + mouldquantitytextBox.Text + "', mould_rate , mould_price, mould_received) VALUES ('" + moulddateTimePicker.Value.ToShortDateString() + "', upper('" + mouldnametextBox.Text + "'), upper('" + mouldproductnametextBox.Text + "'), upper('" + moulddepartmenttextBox.Text + "'), '" + mouldquantitytextBox.Text + "', '" + mouldratetextBox.Text + "', '" + mouldpricetextBox.Text + "', '" + mouldreceivedtextBox.Text + "');", connection);
                        cmd1.ExecuteNonQuery();
                        MessageBox.Show("Record inserted successfully.", "Success", MessageBoxButtons.OK);
                        Clearcontrols();
                    }*/
                }
                catch (Exception ex)
                {
                    MessageBox.Show("There is an error with this insertion! \n\n" + ex, "Alert", MessageBoxButtons.OK);
                }
            }

        }
        private void mouldreportbutton_Click(object sender, EventArgs e)
        {
            SqlCommand tablequery = new SqlCommand("SET DATEFORMAT dmy; select date_received, mould_name, product_name, department, mould_quantity,  mould_rate, total_rate, mould_price, mould_received, received_total from mould_table order by mould_name asc;", connection);
            ConnectionCheck();
            SqlDataReader rdr = tablequery.ExecuteReader();

            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Result";
            sheet.Range["A1"].Value = "Date Received";
            sheet.Range["B1"].Value = "Mould Name";
            sheet.Range["C1"].Value = "Product Name";
            sheet.Range["D1"].Value = "Department";
            sheet.Range["E1"].Value = "Mould Quantity";
            sheet.Range["F1"].Value = "Mould Rate";
            sheet.Range["G1"].Value = "Total Rate";
            sheet.Range["H1"].Value = "Mould Price";
            sheet.Range["I1"].Value = "Mould Received";
            sheet.Range["J1"].Value = "Received Total";
            sheet.Range["A1:J1"].Style.Font.IsBold = true;
            int a = 2;
            while (rdr.Read())
            {
                sheet.Range["A" + a].Value = rdr["date_received"].ToString();
                sheet.Range["B" + a].Value = rdr["mould_name"].ToString();
                sheet.Range["C" + a].Value = rdr["product_name"].ToString();
                sheet.Range["D" + a].Value = rdr["department"].ToString();
                sheet.Range["E" + a].Value = rdr["mould_quantity"].ToString();
                sheet.Range["F" + a].Value = rdr["mould_rate"].ToString();
                sheet.Range["G" + a].Value = rdr["total_rate"].ToString();
                sheet.Range["H" + a].Value = rdr["mould_price"].ToString();
                sheet.Range["I" + a].Value = rdr["mould_received"].ToString();
                sheet.Range["J" + a].Value = rdr["received_total"].ToString();
                a++;
            }
            sheet.AllocatedRange.AutoFitColumns();
            sheet.AllocatedRange.AutoFitRows();
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");
            workbook.SaveToFile("Mould Details Report.xlsx", ExcelVersion.Version2013);


            workbook1 = xlApp.Workbooks.Open(programpath + @"\Mould Details Report.xlsx");
            xlApp.DisplayAlerts = false;
            ((Excel.Worksheet)workbook1.Sheets[2]).Delete();
            workbook1.Save();
            workbook1.Close();
            Marshal.ReleaseComObject(workbook1);
            System.Diagnostics.Process.Start("Mould Details Report.xlsx");
        }

        private void mouldbulkuploadbutton_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Select Mould Excel File";
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                int recordscounta = 0;
                ArrayList mouldids = new ArrayList();
                try
                {
                    Workbook workbook = new Workbook();
                    string filename = openFileDialog1.FileName.ToString();
                    workbook.LoadFromFile(filename);
                    Worksheet sheet = workbook.Worksheets[0];
                    //string a = sheet.Range["A1"].Text;
                    int columncount = Convert.ToInt32(sheet.Columns.LongCount().ToString());
                    int rowcount = Convert.ToInt32(sheet.Rows.LongCount().ToString());
                    int mouldid = 0;
                    //MessageBox.Show("My file is "+ filename +" with cell A1 as " +a+" and with " +columncount +" used colimns and " + rowcount + " used rows.", "yoo", MessageBoxButtons.OK);
                    DialogResult request = MessageBox.Show("Are you sure about the selected file? \n" + filename, "Confirm!", MessageBoxButtons.YesNo);
                    if (request == DialogResult.Yes)
                    {
                        if (columncount == 10 && rowcount > 1)
                        {
                            for (int a = 2; a <= rowcount; a++)
                            {
                                DateTime date = Convert.ToDateTime(sheet.Range["A" + a].Value);
                                string mould_name = sheet.Range["B" + a].Value;
                                string product_name = sheet.Range["C" + a].Value;
                                string department = sheet.Range["D" + a].Value;
                                double quantity = Convert.ToDouble(sheet.Range["E" + a].Value.ToString());
                                double rate = Convert.ToDouble(sheet.Range["F" + a].Value.ToString());
                                //double rate_total = Convert.ToDouble(sheet.Range["G" + a].Text.ToString());
                                double mould_price = Convert.ToDouble(sheet.Range["H" + a].Value.ToString());
                                double received = Convert.ToDouble(sheet.Range["I" + a].Value.ToString());
                                //double received_total = Convert.ToDouble(sheet.Range["J" + a].Value.ToString());
                                ConnectionCheck();
                                SqlCommand cmd1 = new SqlCommand("SET DATEFORMAT dmy; Insert into mould_table (date_received, mould_name, product_name, department, mould_quantity, mould_rate, mould_price, mould_received) VALUES ('" + date + "', upper('" + mould_name + "'), upper('" + product_name + "'), upper('" + department + "'), '" + quantity + "', '" + rate + "', '" + mould_price + "', '" + received + "'); Select MAX(mould_id) from mould_table;", connection);
                                mouldid = (Int32)cmd1.ExecuteScalar();
                                mouldids.Add(mouldid);
                                recordscounta++;
                            }
                            if (recordscounta > 0)
                            { MessageBox.Show(recordscounta + " Record(s) inserted successfully.", "Success", MessageBoxButtons.OK); }
                        }
                        else
                        {
                            MessageBox.Show("Your workbook has more than needed columns, \n Please edit and resubmit", "Error");
                        }
                    }
                }
                catch (Exception ex)
                {
                    int abc = mouldids.Count;
                    string abc123 = String.Join("', '", mouldids.ToArray());
                    ConnectionCheck();
                    SqlCommand delcmd = new SqlCommand("Delete from mould_table where mould_id in ('" + abc123 + "');", connection);
                    delcmd.ExecuteNonQuery();
                    MessageBox.Show("The error is below \n\n" + ex, "error", MessageBoxButtons.OK);
                }
            }
            this.mould_tableTableAdapter.Fill(this.mould_dbDataSet.Mould_table);
        }
        private void generate_template_click()
        {
            Excel.Workbook workbook;
            workbook = xlApp.Workbooks.Add(misValue);
            Excel.Worksheet rm_receipt_sheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);
            Excel.Worksheet rm_discard_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet rm_product_allocation_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet product_receipt_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet product_sale_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet product_miscellaneous_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet product_damage_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet credit_deposit_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet expense_sheet = (Excel.Worksheet)workbook.Worksheets.Add();
            Excel.Worksheet mould_sheet = (Excel.Worksheet)workbook.Worksheets.Add();

            rm_receipt_sheet.Name = "Raw Material Receipt";
            rm_receipt_sheet.Range["A1"].Value = "Date of Receipt";
            rm_receipt_sheet.Range["B1"].Value = "Raw Material Name";
            rm_receipt_sheet.Range["C1"].Value = "Unit of Measurement";
            rm_receipt_sheet.Range["D1"].Value = "Raw Material Description";
            rm_receipt_sheet.Range["E1"].Value = "Unit Cost";
            rm_receipt_sheet.Range["F1"].Value = "Quantity Purchased";
            rm_receipt_sheet.Range["A1:F1"].EntireRow.Font.Bold = true;
            string rm_receipt_sheet_query = "SET DATEFORMAT dmy; Select format(A.date_of_purchase, 'dd-MMM-yyyy') date, B.raw_material_name, B.unit_of_measurement, B.raw_material_desc, A.unit_cost, A.quantity_purchased from raw_material_purchase_table A, raw_material_table B where A.rm_id = B.rm_id;";
            int count1 = 2;
            ConnectionCheck();
            SqlCommand com1 = new SqlCommand(rm_receipt_sheet_query, connection);
            SqlDataReader rdr1 = com1.ExecuteReader();
            while(rdr1.Read())
            {
                rm_receipt_sheet.Range["A" + count1].Value = rdr1["date"].ToString();
                rm_receipt_sheet.Range["B" + count1].Value = rdr1["raw_material_name"].ToString();
                rm_receipt_sheet.Range["C" + count1].Value = rdr1["unit_of_measurement"].ToString();
                rm_receipt_sheet.Range["D" + count1].Value = rdr1["raw_material_desc"].ToString();
                rm_receipt_sheet.Range["E" + count1].Value = rdr1["unit_cost"].ToString();
                rm_receipt_sheet.Range["F" + count1].Value = rdr1["quantity_purchased"].ToString();
                count1++;
            }


            rm_discard_sheet.Name = "Raw Material Discard";
            rm_discard_sheet.Range["A1"].Value = "Date Discard";
            rm_discard_sheet.Range["B1"].Value = "Raw Material Name";
            rm_discard_sheet.Range["C1"].Value = "Quantitiy Discarded";
            rm_discard_sheet.Range["D1"].Value = "Reason";
            rm_discard_sheet.Range["A1:D1"].EntireRow.Font.Bold = true;
            string rm_discard_sheet_query = "SET DATEFORMAT dmy; Select format(A.date, 'dd-MMM-yyyy') date, B.raw_material_name, A.quantity_discarded, A.reason from raw_material_discarded_table A, raw_material_table B where A.rm_id = B.rm_id;";
            int count2 = 2;
            ConnectionCheck();
            SqlCommand com2 = new SqlCommand(rm_discard_sheet_query, connection);
            SqlDataReader rdr2 = com2.ExecuteReader();
            while (rdr2.Read())
            {
                rm_discard_sheet.Range["A" + count2].Value = rdr2["date"].ToString();
                rm_discard_sheet.Range["B" + count2].Value = rdr2["raw_material_name"].ToString();
                rm_discard_sheet.Range["C" + count2].Value = rdr2["quantity_discarded"].ToString();
                rm_discard_sheet.Range["D" + count2].Value = rdr2["reason"].ToString();
                count2++;
            }


            rm_product_allocation_sheet.Name = "Raw Material Prouct Allocation";
            rm_product_allocation_sheet.Range["A1"].Value = "Raw Material Name";
            rm_product_allocation_sheet.Range["B1"].Value = "Product Name";
            rm_product_allocation_sheet.Range["C1"].Value = "Quantity Utilized";
            rm_product_allocation_sheet.Range["A1:C1"].EntireRow.Font.Bold = true;
            string rm_product_allocation_sheet_query = "SET DATEFORMAT dmy; Select  B.raw_material_name, A.product_name, A.rm_unit_quantity_utilized from raw_material_product_allocation_table A, raw_material_table B where A.rm_id = B.rm_id;";
            int count3 = 2;
            ConnectionCheck();
            SqlCommand com3 = new SqlCommand(rm_product_allocation_sheet_query, connection);
            SqlDataReader rdr3 = com3.ExecuteReader();
            while (rdr3.Read())
            {
                rm_product_allocation_sheet.Range["A" + count3].Value = rdr3["raw_material_name"].ToString();
                rm_product_allocation_sheet.Range["B" + count3].Value = rdr3["product_name"].ToString();
                rm_product_allocation_sheet.Range["C" + count3].Value = rdr3["quantity_utilized"].ToString();
                count3++;
            }


            product_receipt_sheet.Name = "Product Receipt";
            product_receipt_sheet.Range["A1"].Value = "Date of Receipt";
            product_receipt_sheet.Range["B1"].Value = "Product Name";
            product_receipt_sheet.Range["C1"].Value = "Unit of Measurement";
            product_receipt_sheet.Range["D1"].Value = "Product Description";
            product_receipt_sheet.Range["E1"].Value = "Current Unit Price";
            product_receipt_sheet.Range["F1"].Value = "Fine Quantity Received";
            product_receipt_sheet.Range["G1"].Value = "Damaged Quantity Received";
            product_receipt_sheet.Range["A1:G1"].EntireRow.Font.Bold = true;
            string product_receipt_sheet_query = "SET DATEFORMAT dmy; Select format(A.date_of_receipt, ('dd-MMM-yyyy')) date, B.product_name, B.unit_of_measurement, B.product_desc, B.current_unit_price, A.fine_quantity_received, A.damaged_quantity_received from product_received_table A, product_table B where A.product_id = B.product_id";
            int count4 = 2;
            ConnectionCheck();
            SqlCommand com4 = new SqlCommand(product_receipt_sheet_query, connection);
            SqlDataReader rdr4 = com4.ExecuteReader();
            while (rdr4.Read())
            {
                product_receipt_sheet.Range["A" + count4].Value = rdr4["date"].ToString();
                product_receipt_sheet.Range["B" + count4].Value = rdr4["product_name"].ToString();
                product_receipt_sheet.Range["C" + count4].Value = rdr4["unit_of_measurement"].ToString();
                product_receipt_sheet.Range["D" + count4].Value = rdr4["product_desc"].ToString();
                product_receipt_sheet.Range["E" + count4].Value = rdr4["current_unit_price"].ToString();
                product_receipt_sheet.Range["F" + count4].Value = rdr4["fine_quantity_received"].ToString();
                product_receipt_sheet.Range["G" + count4].Value = rdr4["damaged_quantity_received"].ToString();
                count4++;
            }

            product_sale_sheet.Name = "Product Sale";
            product_sale_sheet.Range["A1"].Value = "Date of Sale";
            product_sale_sheet.Range["B1"].Value = "Product Name";
            product_sale_sheet.Range["C1"].Value = "Voucher ID";
            product_sale_sheet.Range["D1"].Value = "Payment Type";
            product_sale_sheet.Range["E1"].Value = "Fine Quantity Sold";
            product_sale_sheet.Range["F1"].Value = "Fine Quantity Unit Price";
            product_sale_sheet.Range["G1"].Value = "Damaged Quantity Sold";
            product_sale_sheet.Range["H1"].Value = "Damaged Quantity Unit Price";
            product_sale_sheet.Range["I1"].Value = "Customer Name";
            product_sale_sheet.Range["A1:I1"].EntireRow.Font.Bold = true;
            string product_sale_sheet_query = "SET DATEFORMAT dmy; Select format(A.date_of_sale, ('dd-MMM-yyyy')) date, B.product_name, A.voucher_id,  A.payment_type, A.fine_quantity_sold, A.fine_unit_price, A.damaged_quantity_sold, A.damaged_unit_price, A.customer_name from product_sales_table A, product_table B where A.product_id = B.product_id";
            int count5 = 2;
            ConnectionCheck();
            SqlCommand com5 = new SqlCommand(product_sale_sheet_query, connection);
            SqlDataReader rdr5 = com5.ExecuteReader();
            while (rdr5.Read())
            {
                product_sale_sheet.Range["A" + count5].Value = rdr5["date"].ToString();
                product_sale_sheet.Range["B" + count5].Value = rdr5["product_name"].ToString();
                product_sale_sheet.Range["C" + count5].Value = rdr5["voucher_id"].ToString();
                product_sale_sheet.Range["D" + count5].Value = rdr5["payment_type"].ToString();
                product_sale_sheet.Range["E" + count5].Value = rdr5["fine_quantity_sold"].ToString();
                product_sale_sheet.Range["F" + count5].Value = rdr5["fine_unit_price"].ToString();
                product_sale_sheet.Range["G" + count5].Value = rdr5["damaged_quantity_sold"].ToString();
                product_sale_sheet.Range["H" + count5].Value = rdr5["damaged_unit_price"].ToString();
                product_sale_sheet.Range["I" + count5].Value = rdr5["customer_name"].ToString();
                count5++;
            }


            product_miscellaneous_sheet.Name = "Product Miscellaneous";
            product_miscellaneous_sheet.Range["A1"].Value = "Date Recorded";
            product_miscellaneous_sheet.Range["B1"].Value = "Product Name";
            product_miscellaneous_sheet.Range["C1"].Value = "Fine Quantity";
            product_miscellaneous_sheet.Range["D1"].Value = "Damaged Quantity";
            product_miscellaneous_sheet.Range["E1"].Value = "Reason";
            product_miscellaneous_sheet.Range["A1:E1"].EntireRow.Font.Bold = true;
            string product_miscellaneous_sheet_query = "SET DATEFORMAT dmy; Select format(A.date, ('dd-MMM-yyyy')) date, B.product_name, A.fine_quantity, A.damaged_quantity, A.reason from product_miscellaneous_table A, product_table B where A.product_id = B.product_id";
            int count6 = 2;
            ConnectionCheck();
            SqlCommand com6 = new SqlCommand(product_miscellaneous_sheet_query, connection);
            SqlDataReader rdr6 = com6.ExecuteReader();
            while (rdr6.Read())
            {
                product_miscellaneous_sheet.Range["A" + count6].Value = rdr6["date"].ToString();
                product_miscellaneous_sheet.Range["B" + count6].Value = rdr6["product_name"].ToString();
                product_miscellaneous_sheet.Range["C" + count6].Value = rdr6["fine_quantity"].ToString();
                product_miscellaneous_sheet.Range["D" + count6].Value = rdr6["damaged_quantity"].ToString();
                product_miscellaneous_sheet.Range["E" + count6].Value = rdr6["reason"].ToString();
                count6++;
            }


            product_damage_sheet.Name = "Product Damage";
            product_damage_sheet.Range["A1"].Value = "Date Recorded";
            product_damage_sheet.Range["B1"].Value = "Product Name";
            product_damage_sheet.Range["C1"].Value = "Quantitiy Recorded";
            product_damage_sheet.Range["A1:C1"].EntireRow.Font.Bold = true;
            string product_damage_sheet_query = "SET DATEFORMAT dmy; Select format(A.date, ('dd-MMM-yyyy')) date, B.product_name, A.quantity_damaged from product_damage_table A, product_table B where A.product_id = B.product_id";
            int count7 = 2;
            ConnectionCheck();
            SqlCommand com7 = new SqlCommand(product_damage_sheet_query, connection);
            SqlDataReader rdr7 = com7.ExecuteReader();
            while (rdr7.Read())
            {
                product_damage_sheet.Range["A" + count7].Value = rdr7["date"].ToString();
                product_damage_sheet.Range["B" + count7].Value = rdr7["product_name"].ToString();
                product_damage_sheet.Range["C" + count7].Value = rdr7["quantity_damaged"].ToString();
                count7++;
            }


            credit_deposit_sheet.Name = "Credit Deposit Sheet";
            credit_deposit_sheet.Range["A1"].Value = "Date of Payment";
            credit_deposit_sheet.Range["B1"].Value = "Customer Name";
            credit_deposit_sheet.Range["C1"].Value = "Receipt ID";
            credit_deposit_sheet.Range["D1"].Value = "Payment Amount";
            credit_deposit_sheet.Range["A1:D1"].EntireRow.Font.Bold = true;
            string credit_deposit_sheet_query = "SET DATEFORMAT dmy; Select format(A.date_of_deposit, ('dd-MMM-yyyy')) date, A.customer_name, A.receipt_id, A.amount_paid from product_credit_deposit_table A";
            int count8 = 2;
            ConnectionCheck();
            SqlCommand com8 = new SqlCommand(credit_deposit_sheet_query, connection);
            SqlDataReader rdr8 = com8.ExecuteReader();
            while (rdr8.Read())
            {
                credit_deposit_sheet.Range["A" + count8].Value = rdr8["date"].ToString();
                credit_deposit_sheet.Range["B" + count8].Value = rdr8["customer_name"].ToString();
                credit_deposit_sheet.Range["C" + count8].Value = rdr8["receipt_id"].ToString();
                credit_deposit_sheet.Range["D" + count8].Value = rdr8["amount_paid"].ToString();
                count8++;
            }


            expense_sheet.Name = "Expenses";
            expense_sheet.Range["A1"].Value = "Date Recorded";
            expense_sheet.Range["B1"].Value = "Expense Name";
            expense_sheet.Range["C1"].Value = "Expense Type";
            expense_sheet.Range["D1"].Value = "Expense Detail";
            expense_sheet.Range["E1"].Value = "Amount";
            expense_sheet.Range["A1:E1"].EntireRow.Font.Bold = true;
            string expense_sheet_query = "SET DATEFORMAT dmy; Select format(A.date, ('dd-MMM-yyyy')) date, A.expense_name, A.expense_type, A.expense_detail, A.amount from expenses_table A";
            int count9 = 2;
            ConnectionCheck();
            SqlCommand com9 = new SqlCommand(expense_sheet_query, connection);
            SqlDataReader rdr9 = com9.ExecuteReader();
            while (rdr9.Read())
            {
                expense_sheet.Range["A" + count9].Value = rdr9["date"].ToString();
                expense_sheet.Range["B" + count9].Value = rdr9["expense_name"].ToString();
                expense_sheet.Range["C" + count9].Value = rdr9["expense_type"].ToString();
                expense_sheet.Range["D" + count9].Value = rdr9["expense_detail"].ToString();
                expense_sheet.Range["E" + count9].Value = rdr9["amount"].ToString();
                count9++;
            }


            mould_sheet.Name = "Mould";
            mould_sheet.Range["A1"].Value = "Date Received";
            mould_sheet.Range["B1"].Value = "Mould Name";
            mould_sheet.Range["C1"].Value = "Product Name";
            mould_sheet.Range["D1"].Value = "Department";
            mould_sheet.Range["E1"].Value = "Mould Quantity";
            mould_sheet.Range["F1"].Value = "Mould Rate";
            mould_sheet.Range["G1"].Value = "Mould Price";
            mould_sheet.Range["H1"].Value = "Mould Received";
            mould_sheet.Range["A1:H1"].EntireRow.Font.Bold = true;

            string mould_sheet_query = "SET DATEFORMAT dmy; Select format(A.date_received, ('dd-MMM-yyyy')) date, A.mould_name, A.product_name, A.department, A.mould_quantity, A.mould_rate, A.mould_price, A.mould_received from Mould_table A";
            int count10 = 2;
            ConnectionCheck();
            SqlCommand com10 = new SqlCommand(mould_sheet_query, connection);
            SqlDataReader rdr10 = com10.ExecuteReader();
            while (rdr10.Read())
            {
                mould_sheet.Range["A" + count10].Value = rdr10["date"].ToString();
                mould_sheet.Range["B" + count10].Value = rdr10["mould_name"].ToString();
                mould_sheet.Range["C" + count10].Value = rdr10["product_name"].ToString();
                mould_sheet.Range["D" + count10].Value = rdr10["department"].ToString();
                mould_sheet.Range["E" + count10].Value = rdr10["mould_quantity"].ToString();
                mould_sheet.Range["F" + count10].Value = rdr10["mould_rate"].ToString();
                mould_sheet.Range["G" + count10].Value = rdr10["mould_price"].ToString();
                mould_sheet.Range["H" + count10].Value = rdr10["mould_received"].ToString();
                count10++;
            }


            rm_receipt_sheet.Columns.AutoFit();
            rm_discard_sheet.Columns.AutoFit();
            rm_product_allocation_sheet.Columns.AutoFit();
            product_receipt_sheet.Columns.AutoFit();
            product_sale_sheet.Columns.AutoFit();
            product_miscellaneous_sheet.Columns.AutoFit();
            product_damage_sheet.Columns.AutoFit();
            credit_deposit_sheet.Columns.AutoFit();
            expense_sheet.Columns.AutoFit();
            mould_sheet.Columns.AutoFit();

            rm_receipt_sheet.Move(workbook.Sheets[1]);
            rm_discard_sheet.Move(workbook.Sheets[2]);
            rm_product_allocation_sheet.Move(workbook.Sheets[3]);
            product_receipt_sheet.Move(workbook.Sheets[4]);
            product_sale_sheet.Move(workbook.Sheets[5]);
            product_miscellaneous_sheet.Move(workbook.Sheets[6]);
            product_damage_sheet.Move(workbook.Sheets[7]);
            credit_deposit_sheet.Move(workbook.Sheets[8]);
            expense_sheet.Move(workbook.Sheets[9]);

            xlApp.DisplayAlerts = false;
            Random random = new Random();
            int num = random.Next(155000, 155999);
            string randa = DateTime.Now.ToString("ddMMyyyy");

            workbook.SaveAs(programpath + @"\Hardware Template_" + randa + "_" + num + ".xlsx");
            workbook.Close();
            xlApp.Quit();

            Marshal.ReleaseComObject(rm_receipt_sheet);
            Marshal.ReleaseComObject(rm_discard_sheet);
            Marshal.ReleaseComObject(rm_product_allocation_sheet);
            Marshal.ReleaseComObject(product_receipt_sheet);
            Marshal.ReleaseComObject(product_sale_sheet);
            Marshal.ReleaseComObject(product_miscellaneous_sheet);
            Marshal.ReleaseComObject(product_damage_sheet);
            Marshal.ReleaseComObject(credit_deposit_sheet);
            Marshal.ReleaseComObject(expense_sheet);
            Marshal.ReleaseComObject(mould_sheet);
            Marshal.ReleaseComObject(workbook);

            System.Diagnostics.Process.Start(programpath + @"\Hardware Template_" + randa + "_" + num + ".xlsx");
        }

        private void templatebutton_Click(object sender, EventArgs e)
        {
            generate_template_click();
        }

        private void insertbulkbutton_Click(object sender, EventArgs e)
        {
            int insertioncounter = 0;
            openFileDialog1.Title = "Select Excel File";
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            ConnectionCheck();
            SqlCommand mycmd = new SqlCommand("SET DATEFORMAT dmy;", connection);
            mycmd.ExecuteNonQuery();
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                string errors = "";
                DialogResult result2 = MessageBox.Show("Are you sure you want to upload file: \n\n" + file, "Confirm", MessageBoxButtons.YesNo);
                int sheetcount1 = 0; int sheetcount2 = 0;
                if(result2 == DialogResult.Yes)
                {
                    try
                    {
                        Excel.Workbook workbook = xlApp.Workbooks.Open(file);
                        foreach (Excel.Worksheet sheet in workbook.Sheets)
                        {
                            if(sheet.Name == "Raw Material Receipt")
                            {   sheetcount1++; }
                            if (sheet.Name == "Raw Material Discard")
                            {   sheetcount1++; }
                            if (sheet.Name == "Raw Material Prouct Allocation")
                            {   sheetcount1++; }
                            if (sheet.Name == "Product Receipt")
                            {   sheetcount1++; }
                            if (sheet.Name == "Product Sale")
                            {   sheetcount1++; }
                            if (sheet.Name == "Product Miscellaneous")
                            {   sheetcount1++; }
                            if (sheet.Name == "Product Damage")
                            {   sheetcount1++; }
                            if (sheet.Name == "Credit Deposit Sheet")
                            {   sheetcount1++; }
                            if (sheet.Name == "Expenses")
                            {   sheetcount1++; }
                            if (sheet.Name == "Mould")
                            {   sheetcount1++; }
                            sheetcount2++;
                        }
                        if (sheetcount1 != 10 || sheetcount2 != 10)
                        { MessageBox.Show("The uploaded file is not of the correct format \n Please re-download the template and populate", "Error", MessageBoxButtons.OK); }
                        else
                        {
                            Excel.Worksheet rm_receipt_sheet = workbook.Sheets["Raw Material Receipt"];
                            Excel.Worksheet rm_discard_sheet = workbook.Sheets["Raw Material Discard"];
                            Excel.Worksheet rm_product_allocation_sheet = workbook.Sheets["Raw Material Prouct Allocation"];
                            Excel.Worksheet product_receipt_sheet = workbook.Sheets["Product Receipt"];
                            Excel.Worksheet product_sale_sheet = workbook.Sheets["Product Sale"];
                            Excel.Worksheet product_miscellaneous_sheet = workbook.Sheets["Product Miscellaneous"];
                            Excel.Worksheet product_damage_sheet = workbook.Sheets["Product Damage"];
                            Excel.Worksheet credit_deposit_sheet = workbook.Sheets["Credit Deposit Sheet"];
                            Excel.Worksheet expense_sheet = workbook.Sheets["Expenses"];
                            Excel.Worksheet mould_sheet = workbook.Sheets["Mould"];

                            int rowcount1 = rm_receipt_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount1 = rm_receipt_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount2 = rm_discard_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount2 = rm_discard_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount3 = rm_product_allocation_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount3 = rm_product_allocation_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount4 = product_receipt_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount4 = product_receipt_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount5 = product_sale_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount5 = product_sale_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount6 = product_miscellaneous_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount6 = product_miscellaneous_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount7 = product_damage_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount7 = product_damage_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount8 = credit_deposit_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount8 = credit_deposit_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount9 = expense_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount9 = expense_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
                            int rowcount10 = mould_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                            int colcount10 = mould_sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

                            
                            if (colcount1 == 6 && rowcount1 > 1)
                            {
                                int rmid = 0;
                                ConnectionCheck();
                                int a = 2;
                                int b = 2;
                                SqlCommand mycmd1 = new SqlCommand("Select count(*) counta from raw_material_purchase_table A, raw_material_table B where A.rm_id = B.rm_id;", connection);
                                SqlDataReader myrdr1 = mycmd1.ExecuteReader();
                                if (myrdr1.Read())
                                {
                                    b = Convert.ToInt32(myrdr1["counta"].ToString());
                                    b++;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                try
                                {
                                    for (a = b; a <= rowcount1; a++)
                                    {
                                        
                                        DateTime date = Convert.ToDateTime(rm_receipt_sheet.Range["A" + a].Value);
                                        string rm_name = rm_receipt_sheet.Range["B" + a].Value;
                                        string unit_of_measurement = rm_receipt_sheet.Range["C" + a].Value;
                                        string rm_description = rm_receipt_sheet.Range["D" + a].Value;
                                        double unit_cost = Convert.ToDouble(rm_receipt_sheet.Range["E" + a].Value.ToString());
                                        double quantity_purchased = Convert.ToDouble(rm_receipt_sheet.Range["F" + a].Value.ToString());
                                        
                                        string rawmaterialpurchasequery;
                                        {
                                            ConnectionCheck();
                                            int counta = 0;
                                            SqlCommand cmd = new SqlCommand("Select count(*) counta from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "')", connection);
                                            SqlDataReader rdr = cmd.ExecuteReader();
                                            while (rdr.Read())
                                            {
                                                counta = Convert.ToInt32(rdr["counta"]);
                                            }
                                            if (counta != 0)
                                            {
                                                rawmaterialpurchasequery = "Select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "');";
                                            }
                                            else
                                            {
                                                rawmaterialpurchasequery = "SET DATEFORMAT dmy; Insert into raw_material_table (date_of_added, raw_material_name, raw_material_desc, unit_of_measurement) VALUES('" + date.ToShortDateString() + "', upper('" + rm_name + "'),'" + rm_description + "','" + quantity_purchased + "');";
                                                rawmaterialpurchasequery += "Select MAX(rm_id) rm_id from raw_material_table;";
                                            }
                                            
                                            SqlCommand query_table = new SqlCommand(rawmaterialpurchasequery, connection);
                                            try
                                            {
                                                ConnectionCheck();
                                                rmid = (Int32)query_table.ExecuteScalar();
                                            }
                                            catch (Exception ex)
                                            {
                                                errors += "There is a problem with insertion from Raw Material Receipt worksheet!  Payment line: " + (a) + " \n" + ex + "\n";

                                            }
                                            string rawmaterialpurchasetablequery = "SET DATEFORMAT dmy; Insert into raw_material_purchase_table (date_of_purchase, rm_id, quantity_purchased, unit_cost) VALUES ('" + date.ToShortDateString() + "','" + rmid + "', '" + quantity_purchased + "','" + unit_cost + "');";
                                            if (rmid != 0)
                                            {
                                                SqlCommand com = new SqlCommand(rawmaterialpurchasetablequery, connection);
                                                try
                                                {
                                                    ConnectionCheck();
                                                    com.ExecuteNonQuery();
                                                    string rmcosttble = "SET DATEFORMAT dmy; IF EXISTS (select * from raw_material_current_cost_table where upper(raw_material_name) = upper('" + rm_name + "')) BEGIN Update raw_material_current_cost_table  set current_unit_cost = '" + unit_cost + "' where upper(raw_material_name) = upper('" + rm_name + "') END;";
                                                    rmcosttble += "IF NOT EXISTS (select * from raw_material_current_cost_table where upper(raw_material_name) = upper('" + rm_name + "')) BEGIN INSERT INTO raw_material_current_cost_table (raw_material_name, current_unit_cost) VALUES (upper('" + rm_name + "'), '" + unit_cost + "') END;";
                                                    SqlCommand rmcmd3 = new SqlCommand(rmcosttble, connection);
                                                    ConnectionCheck();
                                                    rmcmd3.ExecuteNonQuery();                                                    
                                                }
                                                catch (Exception ex)
                                                {
                                                    errors += "Failure to update RM purchase table! Payment line: " + (a) + " \n" + ex + "\n";
                                                }
                                            }
                                        }
                                        insertioncounter++;
                                    }
                                }
                                catch(Exception ex)
                                {
                                    errors += "There is a problem with insertion from Raw Material Receipt worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                }
                            }

                            if(colcount2 == 4 && rowcount2 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd2 = new SqlCommand("Select count(*) counta from raw_material_discarded_table A, raw_material_table B where A.rm_id = B.rm_id;", connection);
                                SqlDataReader myrdr2 = mycmd2.ExecuteReader();
                                if (myrdr2.Read())
                                {
                                    b = Convert.ToInt32(myrdr2["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                try
                                {
                                    for (a = b; a <= rowcount2; a++)
                                    {
                                        DateTime date = Convert.ToDateTime(rm_discard_sheet.Range["A" + a].Value);
                                        string rm_name = rm_discard_sheet.Range["B" + a].Value;
                                        double quantity = Convert.ToDouble(rm_discard_sheet.Range["C" + a].Value);
                                        string reason = rm_discard_sheet.Range["D" + a].Value;
                                        string rmdiscardtablequery = "SET DATEFORMAT dmy; Insert into raw_material_discarded_table (date, rm_id, quantity_discarded, unit_cost, reason) VALUES ('" + date.ToShortDateString() + "',(Select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "'),'" + quantity + "', (Select current_unit_cost from raw_material_current_cost_table where upper(raw_material_name) = upper('" + rm_name + "') , '" + reason + "');";
                                        SqlCommand com = new SqlCommand(rmdiscardtablequery, connection);
                                        ConnectionCheck();
                                        com.ExecuteNonQuery();
                                        insertioncounter++;
                                    }
                                }
                                catch(Exception ex)
                                {
                                    errors += "There is a problem with insertion from Raw Material Discard worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                }
                            }

                            if(colcount3 == 3 && rowcount3 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd3 = new SqlCommand("Select count(*) counta from raw_material_product_allocation_table A, raw_material_table B where A.rm_id = B.rm_id;", connection);
                                SqlDataReader myrdr3 = mycmd3.ExecuteReader();
                                if (myrdr3.Read())
                                {
                                    b = Convert.ToInt32(myrdr3["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount3; a++)
                                {
                                    string rmproductallocationclearbuttonquery;
                                    try
                                    {
                                        string rm_name = rm_product_allocation_sheet.Range["A" + a].Value;
                                        string product_name = rm_product_allocation_sheet.Range["B" + a].Value;
                                        double quantity_utilized = Convert.ToDouble(rm_product_allocation_sheet.Range["C" + a].Value);
                                        rmproductallocationclearbuttonquery = "SET DATEFORMAT dmy; IF EXISTS (SELECT * FROM raw_material_product_allocation_table WHERE upper(product_name) = upper('" + product_name + "') and rm_id = (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "')) BEGIN Update   raw_material_product_allocation_table  set  rm_unit_quantity_utilized = '" + quantity_utilized + "' where upper(product_name) =  upper('" + product_name + "' and rm_id = (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "') END;";
                                        rmproductallocationclearbuttonquery += "IF NOT EXISTS (SELECT * FROM raw_material_product_allocation_table WHERE upper(product_name) = upper('" + product_name + "') and rm_id = (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "'))  BEGIN Insert into raw_material_product_allocation_table (product_name, rm_id, rm_unit_quantity_utilized) VALUES (upper('" + product_name + "'), (select rm_id from raw_material_table where upper(raw_material_name) = upper('" + rm_name + "')), '" + quantity_utilized + "') END;";
                                        rmproductallocationclearbuttonquery += "IF NOT EXISTS(SELECT product_name from raw_m_temp_table where upper(product_name) = upper('" + product_name + "') ) BEGIN Insert into raw_m_temp_table (product_name) values (upper('" + product_name + "')) END;";
                                        ConnectionCheck();
                                        SqlCommand cmd4 = new SqlCommand(rmproductallocationclearbuttonquery, connection);
                                        cmd4.ExecuteNonQuery();
                                    }
                                    catch(Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Raw Material Product Allocation worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }

                            if(colcount4 == 7 && rowcount4 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                SqlCommand mycmd4 = new SqlCommand("Select count(*) counta from product_received_table A, product_table B where A.product_id = B.product_id;", connection);
                                SqlDataReader myrdr4 = mycmd4.ExecuteReader();
                                if (myrdr4.Read())
                                {
                                    b = Convert.ToInt32(myrdr4["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount4; a++)
                                {
                                    string productreceiptinsertquery;
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(product_receipt_sheet.Range["A" + a].Value);
                                        string product_name = product_receipt_sheet.Range["B" + a].Value;
                                        string unit_of_measurement = product_receipt_sheet.Range["C" + a].Value;
                                        string product_desc = product_receipt_sheet.Range["D" + a].Value;
                                        double current_unit_price = Convert.ToDouble(product_receipt_sheet.Range["E" + a].Value);
                                        double fine_quantity_received = Convert.ToDouble(product_receipt_sheet.Range["F" + a].Value);
                                        double damage_quantity_received = Convert.ToDouble(product_receipt_sheet.Range["G" + a].Value);
                                        double productunitcost;
                                        int productid;

                                        SqlCommand unitcostquery = new SqlCommand("SET DATEFORMAT dmy; select sum(A.product_unit_cost) unit_cost from (Select A.rm_id, A.product_id product_id, (A.rm_unit_quantity_utilized*B.unit_cost) product_unit_cost from Raw_matrial_product_allocation_table as A, Raw_material_purchase_table as B where A.rm_id = B.rm_id)A inner join Product_table AS B on upper(A.product_name) = upper(B.product_name) where upper(B.product_name) = upper('" + product_name + "') group by A.product_id;", connection);
                                        SqlDataReader rdr2 = unitcostquery.ExecuteReader();
                                        if (rdr2.Read())
                                        { productunitcost = Convert.ToDouble(rdr2["unit_cost"].ToString()); }
                                        else { productunitcost = 0; }
                                        productreceiptinsertquery = "SET DATEFORMAT dmy; IF EXISTS Select * from product_table where upper(product_name) = upper('" + product_name + "') BEGIN Update product_table set current_unit_price = '" + current_unit_price + "', current_unit_cost = '" + productunitcost + "' where upper(product_name) = upper('" + product_name + "'); Select product_id from product_table where upper(product_name) = upper('" + product_name + "'); END";
                                        productreceiptinsertquery += "SET DATEFORMAT dmy; IF NOT EXISTS Select * from product_table where upper(product_name) = upper('" + product_name + "') BEGIN Insert into product_table (date_added,product_name, product_desc, unit_of_measurement, current_unit_price, current_unit_cost) VALUES('" + date.ToShortDateString() + "', '" + product_name + "','" + product_desc + "','" + unit_of_measurement + "','" + current_unit_price + "', '" + productunitcost + "'); Select MAX(product_id) product_id from product_table; END";
                                        ConnectionCheck();
                                        SqlCommand query_table = new SqlCommand(productreceiptinsertquery, connection);
                                        productid = (Int32)query_table.ExecuteScalar();


                                        string productreceivedtablequery = "SET DATEFORMAT dmy; Insert into product_received_table (date_of_receipt, product_id, fine_quantity_received, damaged_quantity_received,unit_cost) VALUES ('" + date.ToShortDateString() + "','" + productid + "', '" + fine_quantity_received + "','" + damage_quantity_received + "','" + productunitcost + "');";
                                        if ( damage_quantity_received != 0)
                                        {
                                            productreceivedtablequery += "SET DATEFORMAT dmy; Insert into product_damage_table (date, product_id, period_of_damage, quantity_damaged, unit_cost) VALUES('" + date.ToShortDateString() + "','" + productid + "','On Receipt','" + damage_quantity_received + "', '" + productunitcost + "');";
                                        }
                                        if (productid != 0)
                                        {
                                            SqlCommand com = new SqlCommand(productreceivedtablequery, connection);
                                            ConnectionCheck();
                                            com.ExecuteNonQuery();
                                            int abc = 0;
                                            ArrayList rmidlist = new ArrayList();
                                            SqlCommand rmcmd5 = new SqlCommand("Select rm_id from raw_material_product_allocation_table where upper(product_name) = upper ('" + product_name + "');", connection);
                                            ConnectionCheck();
                                            SqlDataReader rmrdr5 = rmcmd5.ExecuteReader();
                                            while (rmrdr5.Read())
                                            {
                                                rmidlist.Add(rmrdr5["rm_id"].ToString());
                                            }
                                            abc = rmidlist.Count;
                                            for (int d = 0; d < abc; d++)
                                            {
                                                string resultspool = "SET DATEFORMAT dmy; Select A.rm_id, A.rm_unit_quantity_utilized rm_unit_quantity_utilized , B.current_unit_cost current_unit_cost from raw_material_product_allocation_table AS A, raw_material_current_cost_table AS B, raw_material_table AS C where A.rm_id = C.rm_id and upper(B.raw_material_name) = upper(C.raw_material_name) and A.rm_id = '" + rmidlist[d] + "';";
                                                SqlCommand resultcmd = new SqlCommand(resultspool, connection);
                                                SqlDataReader resultrdr = resultcmd.ExecuteReader();
                                                while (resultrdr.Read())
                                                {
                                                    string rmusageinsert = "SET DATEFORMAT dmy; Insert into raw_material_to_product_usage_table (date_used, rm_id, quantity_used, rm_unit_cost, product_id) VALUES('" + date.ToShortDateString() + "', '" + rmidlist[d] + "','" + (Convert.ToDouble(resultrdr["rm_unit_quantity_utilized"].ToString()) * (Convert.ToDouble(fine_quantity_received) + Convert.ToDouble(damage_quantity_received))) + "','" + resultrdr["current_unit_cost"].ToString() + "', produuctid);";
                                                    SqlCommand finacmd = new SqlCommand(rmusageinsert, connection);
                                                    ConnectionCheck();
                                                    finacmd.ExecuteNonQuery();
                                                }
                                            }
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Product Receipt worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }
                            if (colcount5 == 9 && rowcount5 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd5 = new SqlCommand("Select  count(*) counta from product_sales_table A, product_table B where A.product_id = B.product_id;", connection);
                                SqlDataReader myrdr5 = mycmd5.ExecuteReader();
                                if (myrdr5.Read())
                                {
                                    b = Convert.ToInt32(myrdr5["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount5; a++)
                                {
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(product_sale_sheet.Range["A" + a].Value);
                                        string product_name = product_sale_sheet.Range["B" + a].Value;
                                        string voucher_id = product_sale_sheet.Range["C" + a].Value;
                                        string payment_type = product_sale_sheet.Range["D" + a].Value;
                                        double fine_quantity_sold = Convert.ToDouble(product_sale_sheet.Range["E" + a].Value);
                                        double fine_quantity_unit_price = Convert.ToDouble(product_sale_sheet.Range["F" + a].Value);
                                        double damaged_quantity_sold = Convert.ToDouble(product_sale_sheet.Range["G" + a].Value);
                                        double damaged_quantity_unit_price = Convert.ToDouble(product_sale_sheet.Range["H" + a].Value);
                                        double customer_name = Convert.ToDouble(product_sale_sheet.Range["I" + a].Value);
                                        double totalamount = (fine_quantity_sold * fine_quantity_unit_price) + (damaged_quantity_sold * damaged_quantity_unit_price);
                                        double productunitcost;
                                        SqlCommand unitcostquery = new SqlCommand("SET DATEFORMAT dmy; select sum(A.product_unit_cost) unit_cost from (Select A.rm_id, A.product_id product_id, (A.rm_unit_quantity_utilized*B.unit_cost) product_unit_cost from Raw_matrial_product_allocation_table as A, Raw_material_purchase_table as B where A.rm_id = B.rm_id)A inner join Product_table AS B on upper(A.product_name) = upper(B.product_name) where upper(B.product_name) = upper('" + product_name + "') group by A.product_id;", connection);
                                        ConnectionCheck();
                                        SqlDataReader rdr2 = unitcostquery.ExecuteReader();
                                        if (rdr2.Read())
                                        { productunitcost = Convert.ToDouble(rdr2["unit_cost"].ToString()); }
                                        else { productunitcost = 0; }
                                        ConnectionCheck();
                                        string salesinsertsearchquery = "SET DATEFORMAT dmy; Insert into product_sales_table (date_of_sale, product_id, voucher_id, payment_type, fine_quantity_sold, fine_unit_price, damages_quantity_sold, damaged_unit_price, customer_name, unit_cost) VALUES('" + date.ToShortDateString() + "', (Select product_id from product_table where upper(product_name) = upper('" + product_name + "') ), '" + voucher_id + "', '" + payment_type + "', '" + fine_quantity_sold + "','" + fine_quantity_unit_price + "','" + damaged_quantity_sold + "','" + damaged_quantity_unit_price + "','" + customer_name + "', '" + productunitcost + "');";

                                        if (paymenttypecomboBox.SelectedItem.ToString().ToUpper() == "CREDIT")
                                        {
                                            SqlCommand cmd2 = new SqlCommand("Select count(*) from product_customer_credit_table where upper(customer_name) = upper('" + customer_name + "'));", connection);
                                            int product = 0;
                                            ConnectionCheck();
                                            product = (Int32)cmd2.ExecuteScalar();
                                            if (product == 0)
                                            {
                                                salesinsertsearchquery += "SET DATEFORMAT dmy; Insert into product_customer_credit_table (date_of_sale, customer_name, total_credit_amount) VALUES ('" + date.ToShortDateString() + "',upper('" + customer_name + "'),'" + totalamount + "');";
                                            }
                                            else
                                            {
                                                salesinsertsearchquery += "SET DATEFORMAT dmy; Update product_customer_credit_table set total_credit_amount = (total_credit_amount + " + totalamount + ") where upper(customer_name) = upper('" + customer_name + "');";
                                            }

                                            SqlCommand cmd3 = new SqlCommand("Select count(*) from product_credit_table where voucher_id = upper('" + voucher_id + "'));", connection);
                                            int product2 = 0;
                                            product = (Int32)cmd3.ExecuteScalar();
                                            if (product2 == 0)
                                            {
                                                salesinsertsearchquery += "SET DATEFORMAT dmy; Insert into product_credit_table (date_of_sale, voucher_id, customer_name, credit_amount) VALUES ('" + date.ToShortDateString() + "','" + voucher_id + "','" + customer_name + "','" + totalamount + "');";
                                            }
                                            else
                                            {
                                                salesinsertsearchquery += "SET DATEFORMAT dmy; Update product_credit_table set credit_amount = (credit_amount + " + totalamount + ") where voucher_id = '" + voucher_id + "';";
                                            }
                                        }
                                        SqlCommand com = new SqlCommand(salesinsertsearchquery, connection);
                                        ConnectionCheck();
                                        com.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Product Sales worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }

                            if (colcount6 == 5 && rowcount6 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd6 = new SqlCommand("Select  count(*) counta from product_miscellaneous_table A, product_table B where A.product_id = B.product_id;", connection);
                                SqlDataReader myrdr6 = mycmd6.ExecuteReader();
                                if (myrdr6.Read())
                                {
                                    b = Convert.ToInt32(myrdr6["counta"].ToString());
                                    b = b + 2;
                                }
                                if(b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount6; a++)
                                {
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(product_miscellaneous_sheet.Range["A" + a].Value);
                                        string product_name = product_miscellaneous_sheet.Range["B" + a].Value;
                                        double fine_quantity = Convert.ToDouble(product_miscellaneous_sheet.Range["C" + a].Value);
                                        double damaged_quantity = Convert.ToDouble(product_miscellaneous_sheet.Range["D" + a].Value);
                                        string reason = product_miscellaneous_sheet.Range["E" + a].Value;

                                        string miscellaneoustablequery = "SET DATEFORMAT dmy; Insert into product_miscellaneous_table (date, product_id, fine_quantity, damaged_quantity, reason, unit_cost) VALUES ('" + date.ToShortDateString() + "',(Select product_id from product_table where upper(product_name) = upper('" + product_name + "'),'" + fine_quantity + "', '" + damaged_quantity + "', '" + productmiscellaneousreasontextBox.Text + "','" + reason + "');";
                                        SqlCommand com = new SqlCommand(miscellaneoustablequery, connection);
                                        ConnectionCheck();
                                        com.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Product Miscellaneous worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }

                            if (colcount7 == 3 && rowcount7 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd7 = new SqlCommand("Select  count(*) counta from product_damage_table A, product_table B where A.product_id = B.product_id;", connection);
                                SqlDataReader myrdr7 = mycmd7.ExecuteReader();
                                if (myrdr7.Read())
                                {
                                    b = Convert.ToInt32(myrdr7["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount7; a++)
                                {
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(product_damage_sheet.Range["A" + a].Value);
                                        string product_name = product_damage_sheet.Range["B" + a].Value;
                                        double quantity_recorded = Convert.ToDouble(product_damage_sheet.Range["C" + a].Value);
                                        double pdtunitprice = 0;

                                        ConnectionCheck();
                                        SqlCommand cmd = new SqlCommand("SET DATEFORMAT dmy; Select TOP 1 unit_cost from product_received_table where product_id = (select product_id from product_table where upper (product_name) = upper('" + product_name + "')) ORDER BY date_of_receipt DESC;", connection);
                                        SqlDataReader rdr = cmd.ExecuteReader();
                                        while (rdr.Read())
                                        {
                                            pdtunitprice = Convert.ToDouble(rdr["unit_cost"].ToString());
                                        }

                                        string discardtablequery = "SET DATEFORMAT dmy; Insert into product_damage_table (date, product_id, period_of_damage, quantity_damaged, unit_cost) VALUES('" + date.ToShortDateString() + "',(Select product_id from product_table where upper(product_name) = upper('" + product_name + "'),'After Receipt','" + quantity_recorded + "', '" + pdtunitprice + "');";
                                        SqlCommand com = new SqlCommand(discardtablequery, connection);
                                        ConnectionCheck();
                                        com.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Product Damage worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }

                            if (colcount8 == 4 && rowcount8 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd8 = new SqlCommand("Select  count(*) counta from product_credit_deposit_table A;", connection);
                                SqlDataReader myrdr8 = mycmd8.ExecuteReader();
                                if (myrdr8.Read())
                                {
                                    b = Convert.ToInt32(myrdr8["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount8; a++)
                                {
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(credit_deposit_sheet.Range["A" + a].Value);
                                        string customer_name = credit_deposit_sheet.Range["B" + a].Value;
                                        string receipt_id = credit_deposit_sheet.Range["C" + a].Value;
                                        double payment_amount = Convert.ToDouble(credit_deposit_sheet.Range["D" + a].Value);

                                        string creditdepositquery = "SET DATEFORMAT dmy; Insert into product_credit_deposit_table (date_of_deposit, receipt_id, amount_paid, customer_name) VALUES('" + date.ToShortDateString() + "', '" + receipt_id + "','" + payment_amount + "','" + customer_name + "');";
                                        creditdepositquery += "SET DATEFORMAT dmy; update product_customer_credit_table set amount_paid = (amount_paid+" + Convert.ToInt32(payment_amount) + ") where upper(customer_name) = upper('" + customer_name + "'); ";
                                        SqlCommand com = new SqlCommand(creditdepositquery, connection);
                                        ConnectionCheck();
                                        com.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Credit Deposit worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }


                            if (colcount9 == 5 && rowcount9 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd9 = new SqlCommand("Select count(*) counta from expenses_table A;", connection);
                                SqlDataReader myrdr9 = mycmd9.ExecuteReader();
                                if (myrdr9.Read())
                                {
                                    b = Convert.ToInt32(myrdr9["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount9; a++)
                                {

                                    int ac = 0;
                                    string expenseinsertsearchquery = "";

                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(expense_sheet.Range["A" + a].Value);
                                        ac = 1;
                                        string expense_name = Convert.ToString(expense_sheet.Range["B" + a].Value);
                                        ac = 2;
                                        string expense_type = Convert.ToString(expense_sheet.Range["C" + a].Value);
                                        ac = 3;
                                        string expense_detail = Convert.ToString(expense_sheet.Range["D" + a].Value);
                                        ac = 4;
                                        double expense_amount = Convert.ToDouble(expense_sheet.Range["E" + a].Value);
                                        ac = 5;
                                        ConnectionCheck();
                                         expenseinsertsearchquery = "SET DATEFORMAT dmy; Insert into expenses_table (date, expense_name, expense_type, expense_detail, amount) VALUES('" + date.ToShortDateString() + "', '" + expense_name + "', '" + expense_type + "', '" + expense_detail + "', '" + expense_amount + "')";
                                        SqlCommand com = new SqlCommand(expenseinsertsearchquery, connection);
                                        com.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Expenses worksheet! Payment line: " + (a) + " - " + ac + " - " + expenseinsertsearchquery + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }

                            if (colcount10 == 8 && rowcount10 > 1)
                            {
                                int a = 2;
                                int b = 2;
                                ConnectionCheck();
                                SqlCommand mycmd10 = new SqlCommand("Select  count(*) counta from Mould_table A;", connection);
                                SqlDataReader myrdr10 = mycmd10.ExecuteReader();
                                if (myrdr10.Read())
                                {
                                    b = Convert.ToInt32(myrdr10["counta"].ToString());
                                    b = b + 2;
                                }
                                if (b < 2)
                                {
                                    b = 2;
                                }
                                for (a = b; a <= rowcount10; a++)
                                {
                                    try
                                    {
                                        DateTime date = Convert.ToDateTime(mould_sheet.Range["A" + a].Value);
                                        string mould_name = mould_sheet.Range["B" + a].Value;
                                        string product_name = mould_sheet.Range["C" + a].Value;
                                        string department = mould_sheet.Range["D" + a].Value;
                                        double mould_quantity = Convert.ToDouble(mould_sheet.Range["E" + a].Value);
                                        double mould_rate = Convert.ToDouble(mould_sheet.Range["F" + a].Value);
                                        double mould_price = Convert.ToDouble(mould_sheet.Range["G" + a].Value);
                                        double mould_received = Convert.ToDouble(mould_sheet.Range["H" + a].Value);

                                        ConnectionCheck();
                                        SqlCommand cmd1 = new SqlCommand("SET DATEFORMAT dmy; Insert into mould_table (date_received, mould_name, product_name, department, mould_quantity, mould_rate, mould_price, mould_received) VALUES ('" + date.ToShortDateString() + "', upper('" + mould_name + "'), upper('" + product_name + "'), upper('" + department + "'), '" + mould_quantity + "', '" + mould_rate + "', '" + mould_price + "', '" + mould_received + "');", connection);
                                        cmd1.ExecuteNonQuery();
                                    }
                                    catch (Exception ex)
                                    {
                                        errors += "There is a problem with insertion from Mould worksheet! Payment line: " + (a) + " \n" + ex + ". \n";
                                    }
                                    insertioncounter++;
                                }
                            }
                     

                            if(errors != "")
                            {
                                MessageBox.Show(errors, "Error");
                            }
                            else if(insertioncounter == 0)
                            {
                                MessageBox.Show("There is no data to insert!!", "Notification");
                            }
                            else if(insertioncounter > 0)
                            {
                                MessageBox.Show("All Data Inserted Successfully: ", "Success");
                            }

                            this.raw_material_tableTableAdapter.Fill(this.rawmaterialnameDataSet.raw_material_table);
                            this.product_tableTableAdapter.Fill(this.productnamelistDataset.Product_table);
                            this.expenses_tableTableAdapter.Fill(this.expenseDataSet.expenses_table);
                            this.mould_tableTableAdapter.Fill(this.mould_dbDataSet.Mould_table);

                            xlApp.DisplayAlerts = false;
                            Marshal.ReleaseComObject(rm_receipt_sheet);
                            Marshal.ReleaseComObject(rm_discard_sheet);
                            Marshal.ReleaseComObject(rm_product_allocation_sheet);
                            Marshal.ReleaseComObject(product_receipt_sheet);
                            Marshal.ReleaseComObject(product_sale_sheet);
                            Marshal.ReleaseComObject(product_miscellaneous_sheet);
                            Marshal.ReleaseComObject(product_damage_sheet);
                            Marshal.ReleaseComObject(credit_deposit_sheet);
                            Marshal.ReleaseComObject(expense_sheet);
                            Marshal.ReleaseComObject(mould_sheet);
                            Marshal.ReleaseComObject(workbook);
                        }
                    }
                    catch(Exception EX)
                    {
                        MessageBox.Show("Error is below: \n\n" +EX, "Error", MessageBoxButtons.OK);
                    }
                }
            }
        }

        private void rawmaterialpurchaseeditunitcosttextBox_Click(object sender, EventArgs e)
        {
            if(editrawmaterialunitcostcheckBox.Checked == true)
            {
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = false;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
            else
            {
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
        }

        private void editrawmaterialunitcostcheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (editrawmaterialunitcostcheckBox.Checked == true)
            {
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = false;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
            else
            {
                rawmaterialpurchaseeditunitcosttextBox.ReadOnly = true;
                rawmaterialpurchaseeditunitcosttextBox.Text = "";
            }
        }

        private void databasereset_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Are you sure you want \nto empty the DATABASE!? \n", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
            if (result1 == DialogResult.Yes)
            {
                DialogResult result2 = MessageBox.Show("Please Confirm  \nYou are Deleting the DATABASE. \nThis cannot be undone!!", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                if(result2 == DialogResult.Yes)
                {
                    SqlCommand cmd = new SqlCommand("select name from sys.tables where type_desc = 'USER_TABLE'", connection);
                    ArrayList arl = new ArrayList();
                    ConnectionCheck();
                    SqlDataReader rmrdr5 = cmd.ExecuteReader();
                    while (rmrdr5.Read())
                    {
                        arl.Add(rmrdr5["name"].ToString());
                    }
                    foreach(string i in arl)
                    {
                        SqlCommand cmd2 = new SqlCommand("Delete from " + i + ";", connection); ConnectionCheck();
                        cmd2.ExecuteNonQuery();
                    }
                    MessageBox.Show("All is done");
                }
                else
                {
                    MessageBox.Show("Operation Cancelled");
                }
            }
            else
            {
                MessageBox.Show("Operation Cancelled");
            }
        }
    }
}