//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//  <copyright file="SM_copyright.cs" company="Total System Services">
//
//  Copyright © TSYS , 2019. All rights reserved.
//
//  Contributor(s) : Adam Gemperline, DBA , 2019
//            
//  </copyright>  
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Xml;
using System.Data.Common;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using ScriptManagerV4.RMDataSet8TableAdapters;
using Microsoft.VisualBasic;
using System.Timers;


namespace ScriptManagerV4
{
    public partial class MainForm : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        public static Stream myStream;      // used for file stream
        private string actionSelected = ""; // stores user selected action 
        string selectedID = "";
        DataTable dt;
        public static Color Yellow;
        private int currentIndex;           // denotes where a RUN sequence stops and proceeds
        private bool proceed;               // controls continuation or exit of an execution
        private bool ignoreAll;             // ignore all error dialogs for RUN sequence
        private int placeholder;
        private bool lightSwitch = false;   // color scheme switch
        public double progress;             // progressBar value
        private int runEnvironment;         // stores index of Run environment selection (comboBox5)
        public double pp1;
        public int pp2;
        private int totalProcesses;
        private int progressPosition;


        public MainForm()
        {
            InitializeComponent();
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // TODO: This line of code loads data into the 'rMDataSet.refresh_testing' table. You can move, or remove it, as needed.
                this.refresh_testingTableAdapter1.Fill(this.rMDataSet8.refresh_testing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\nTry restarting the application.", "Connection Error");
            }
            setInitState();
        }



        private void setInitState()
        {
            toExecLog("\n------------------------------------------------------------------------------" +
                              "\nInitializing states...");

            // initial/refresh states and values
            progress = 0.0;
            progressPosition = 0;
            totalProcesses = 0;
            textBox1.BackColor = Color.DimGray;
            textBox1.Visible = false;
            textBox2.BackColor = Color.DimGray;
            textBox2.Visible = false;
            textBox3.BackColor = Color.DimGray;
            label1.Visible = false;
            textBox3.Visible = false;
            textBox5.Visible = false;
            label17.Visible = false;
            label10.Visible = false;
            label11.Visible = false;
            label13.Visible = false;
            label14.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            comboBox2.Visible = false;
            comboBox3.Visible = false;
            comboBox5.Visible = false;
            currentIndex = 0;
            placeholder = 0;
            ignoreAll = false;
            textBox1.Text = null;
            textBox1.Enabled = false;
            richTextBox1.Text = null;
            richTextBox1.Enabled = false;
            toolStripButton2.Image = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\icons\lock.png");
            toolStripButton2.ToolTipText = "Unlock editor";
            myStream = null;
            label1.Text = "*select .SQL files only";
            textBox2.Text = null;
            textBox3.Text = null;
            textBox5.Text = null;
            actionSelected = null;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox5.Enabled = false;
            comboBox1.Visible = true;
            label9.Visible = true;
            label10.Enabled = true;
            comboBox2.Enabled = false;
            comboBox3.Enabled = false;
            comboBox3.SelectedIndex = -1;
            textBox3.Enabled = false;
            label13.Enabled = true;
            label12.Visible = true;
            label12.Enabled = true;
            label11.Enabled = true;
            label14.Enabled = true;
            label4.Enabled = true;
            label5.Enabled = true;
            toolStripComboBox1.SelectedIndex = 5;
            toolStripTextBox1.Text = null;
            toolStripProgressBar1.Minimum = 0;
            radioButton1.Visible = false;
            radioButton3.Visible = false;
            radioButton4.Visible = false;
            textBox1.KeyDown += new KeyEventHandler(execute_KeyDown);
            textBox3.KeyDown += new KeyEventHandler(execute_KeyDown);
            toolStripTextBox1.KeyUp += new KeyEventHandler(toolStripTextBox1_KeyUp);
            tabControl2.SelectedIndex = 0;
            toolStripLabel3.Text = "";
            toolStripLabel3.Visible = false;
            radioButton3.Text = "Start At Script Order: ";
            radioButton1.Enabled = true;
            radioButton3.Enabled = true;
            radioButton4.Enabled = true;
            toolStripLabel4.Visible = false;

            if (comboBox1.Items.Contains("UPDATE & RESUME"))
                comboBox1.Items.Remove("UPDATE & RESUME");
            if (!comboBox1.Items.Contains("INSERT INTO"))
                comboBox1.Items.Add("INSERT INTO");
            if (!comboBox1.Items.Contains("RUN"))
                comboBox1.Items.Add("RUN");

            loadData();
            dataGridView1.ClearSelection();

            toExecLog("States initialized");

            // set progress bar defaults after a delay
            delay(3000); // delay in milliseconds
            toolStripProgressBar1.Value = 0;
            toolStripProgressBar1.Visible = false;
            toolStripLabel2.Text = "0%";
            toolStripLabel2.Visible = false;
            toolStripLabel3.Visible = false;
            toolStripProgressBar1.Value = 0;
        }



        private void button5_Click(object sender, EventArgs e)
        {
            // CLOSE button
            Close();
        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // ACTION comboBox
            actionSelected = comboBox1.GetItemText(comboBox1.SelectedItem);

            if (actionSelected == "INSERT INTO")
            {
                toggleFieldVisibility(1);
                clearFields(sender, e);
                label10.Enabled = true;
                label14.Enabled = true;
                label14.Visible = true;
                textBox1.Enabled = true;
                textBox3.Enabled = true;
                label17.Visible = false;
                label12.Text = "INSERT INTO:";
                comboBox2.Enabled = true;
                label11.Enabled = true;
                label13.Enabled = true;
                label5.Enabled = true;
                comboBox5.Visible = false;
                textBox2.Text = "AUTO";
                textBox5.Visible = false;
                radioButton1.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
                comboBox3.Enabled = true;
                textBox1.ReadOnly = false;
                textBox3.ReadOnly = false;
            }
            else if (actionSelected == "UPDATE")
            {
                textBox1.ReadOnly = false;
                textBox3.ReadOnly = false;
                toggleFieldVisibility(1);
                label10.Enabled = true;
                label14.Enabled = true;
                label12.Text = "UPDATE:";
                label14.Visible = true;
                label17.Enabled = true;
                label17.Visible = false;
                comboBox2.Enabled = true;
                label11.Enabled = true;
                textBox3.Enabled = true;
                textBox3.Visible = true;
                label13.Enabled = true;
                comboBox5.Visible = false;
                comboBox3.Enabled = true;
                textBox1.Enabled = true;
                label5.Enabled = true;
                textBox2.Visible = true;
                label11.Visible = true;
                setSelectedValues(sender);
                textBox5.Visible = false;
                radioButton1.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
            }
            else if (actionSelected == "RUN")
            {
                toggleFieldVisibility(0);
                label4.Enabled = false;
                label4.Visible = false;
                label5.Visible = false;
                label5.Enabled = false;
                label10.Enabled = false;
                label10.Visible = false;
                label14.Visible = false;
                label14.Enabled = false;
                textBox1.Enabled = false;
                label17.Enabled = true;
                label17.Visible = true;
                label17.Location = new Point(364, 219);
                label12.Text = "Viewing:";
                comboBox2.Enabled = false;
                comboBox2.Visible = false;
                label11.Enabled = false;
                label11.Visible = false;
                textBox3.Enabled = false;
                textBox3.Visible = true;
                label13.Enabled = false;
                label13.Visible = false;
                comboBox5.Visible = true;
                textBox2.Text = "AUTO";
                label17.Text = "Run Environment:";
                comboBox5.Enabled = true;
                comboBox5.Location = new Point(367, 239);
                comboBox3.Enabled = false;
                comboBox3.Visible = false;
                textBox5.Visible = true;
                radioButton1.Visible = true;
                radioButton3.Visible = true;
                radioButton4.Visible = true;
                radioButton1.Checked = false;
                radioButton3.Checked = false;
                radioButton4.Checked = false;
                textBox1.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                clearFields(sender, e);
            }

            // when UPDATE & RESUME is selected
            else if (actionSelected == "UPDATE & RESUME")
            {
                radioButton3.Text = "Resume At S.O. #: ";
                textBox5.Text = dataGridView1.SelectedRows[0].Cells[0].ToString();
                label10.Enabled = true;
                label14.Enabled = true;
                textBox1.Enabled = true;
                label12.Text = "Viewing:";
                comboBox2.Enabled = true;
                comboBox3.Enabled = true;
                label11.Enabled = true;
                textBox3.Enabled = true;
                label13.Enabled = true;
                textBox2.Text = "AUTO";
                label17.Text = "Run Environment:";
                comboBox5.Enabled = true;
                comboBox5.Visible = true;
                textBox5.Text = currentIndex.ToString();
                radioButton1.Enabled = false;
                radioButton3.Enabled = true;
                radioButton4.Enabled = false;
                radioButton1.Checked = false;
                radioButton3.Checked = true;
                radioButton4.Checked = false;
                radioButton1.Visible = true;
                radioButton3.Visible = true;
                radioButton4.Visible = true;
                richTextBox1.Enabled = true;
                toggleFieldVisibility(1);
                setSelectedValues(sender);
                if (runEnvironment != -1)
                    comboBox5.SelectedIndex = runEnvironment;
            }

            // no selection
            else if (comboBox1.SelectedIndex == -1)
            {
                label12.Text = "Viewing:";
                toggleFieldVisibility(0);
                comboBox5.Visible = false;
                label17.Visible = false;
                radioButton1.Visible = false;
                radioButton3.Visible = false;
                radioButton4.Visible = false;
                textBox5.Visible = false;
            }

            if (label5.Enabled)
                label4.Enabled = true;
            else
                label4.Enabled = false;
        }



        private void button4_Click(object sender, EventArgs e)
        {
            toExecLog("Restarting...");
            {
                if (conn != null && conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                setInitState();
            }
        }



        // START: --------------------------------------------------------- EXECUTION STATEMENTS --------------------------------------------------------------------------------------

        // CAUTION! 

        private void button3_Click(object sender, EventArgs e)
        {
            // EXECUTE button 
            proceed = true;     // control variable that stops RUN sequence upon user's error response
            button3.Enabled = false;
            button4.Enabled = false;
            toolStripProgressBar1.Visible = true;
            toolStripLabel2.Enabled = true;
            toolStripLabel2.Visible = true;
            runEnvironment = comboBox5.SelectedIndex; // store Run environment for when it gets reset by execution processes
            toolStripLabel3.IsLink = false;

            // determine number of records in table
            int numRows = dataGridView1.Rows.Count;
            int lastIndexInView = dataGridView1.Rows.GetLastRow(DataGridViewElementStates.Visible);
            int lastIndex = numRows - 1;

            // show Execution Summary tab
            tabControl1.SelectedIndex = 1;
            tabControl1.Update();

            if (conn.State == ConnectionState.Open)
                conn.Close();

            if (fieldCheck(sender, e) == false)  // checks for required information, returns false if incomplete
            {
                button3.Enabled = true;
                button4.Enabled = true;
                return;
            }

            if (actionSelected == "RUN")
            {
                int startingIndex;

                // if "From Table View" is selected
                if ((radioButton1.Checked == true) && (comboBox5.SelectedIndex != -1))
                {
                    // select first row 
                    dataGridView1.Rows[0].Selected = true;

                    toolStripLabel4.Visible = false;
                    startingIndex = 0;
                    Task task = runMultiple(sender, numRows, lastIndexInView, e, startingIndex);
                }
                // Start/Resume At Script Order is selected 
                else if (radioButton3.Checked == true && (comboBox5.SelectedIndex != -1))
                {
                    toolStripLabel4.Visible = false;

                    if (textBox5.Text != dataGridView1.SelectedRows[0].Cells[0].Value.ToString())
                    {
                        searchForScriptOrder(textBox5.Text);
                    }
                    startingIndex = dataGridView1.SelectedRows[0].Index;
                    Task task = runMultiple(sender, numRows, lastIndexInView, e, startingIndex);
                }
                // Run Selected Script is selected 
                else if ((radioButton4.Checked == true) && (comboBox5.SelectedIndex != -1))
                {
                    toolStripLabel4.Visible = false;
                    toolStripLabel3.Text = "Executing script...";
                    toolStripLabel3.Visible = true;
                    toolStripProgressBar1.Value = 0;
                    toolStripProgressBar1.Maximum = 2;
                    updateProgressBar();
                    DataGridViewRow row = dataGridView1.SelectedRows[0];
                    runScript(sender, row, lastIndexInView, e);
                    updateProgressBar();
                    toolStripLabel3.Text = "Execution Complete";
                    toolStrip3.Update();
                }
            }
            else if (actionSelected == "UPDATE" && proceed == true)
            {
                updateRecord(sender, e);
            }
            else if (actionSelected == "INSERT INTO")
            {
                proceed = false;

                // function returns -1 if no matching script order is found
                int searchResult = searchForScriptOrder(textBox1.Text);

                if (searchResult == -1)
                {
                    toolStripProgressBar1.Maximum = 1;
                    insertIntoTable(sender, e);
                }
                else
                {
                    MessageBox.Show("That Script Order number already exists. Feel free to utilize up to two decimal places.");
                    button3.Enabled = true;
                    button4.Enabled = true;
                    return;
                }
            }
            else if (actionSelected == "UPDATE & RESUME")
            {
                int startingIndex = 0;
                proceed = true;
                updateRecord(sender, e);

                if (!comboBox5.Items.Contains("RUN"))
                    comboBox1.Items.Add("RUN");
                comboBox1.SelectedItem = "RUN";
                radioButton3.Checked = true; // run option: Start/Resume At Selected Script Order

                int index = dataGridView1.SelectedRows[0].Index; // store the current selected record before loadData() is called, as it will be deselected
                loadData();
                dataGridView1.Rows[index].Selected = true; // reselect the record 

                if (dataGridView1.SelectedRows[0] != null && dataGridView1.SelectedRows[0].Index > -1)
                {
                    startingIndex = dataGridView1.SelectedRows[0].Index; // start RUN from selected row's index
                    Task task = runMultiple(sender, numRows, lastIndexInView, e, startingIndex);
                }
                else
                {
                    MessageBox.Show("Please select a starting point");
                }
            }

            button3.Enabled = true;
            button4.Enabled = true;

            if (toolStripProgressBar1.Value == toolStripProgressBar1.Maximum)
            {
                toExecLog("Execution complete");
                toolStripLabel3.Text = "Execution complete";
                // print execution summary
                executionSummary("display");
                // restart
                button4_Click(sender, e);
            }
        }



        // ------------------------------------------- Run all records in table (filter applies) ----------------------------------------
        private void runFromTableView(object sender, int numRows, int lastIndexInView, EventArgs e)
        {
            /*           toolStripProgressBar1.Maximum = numRows;
                       toolStripLabel3.Text = "Executing scripts...";
                       toolStripLabel3.Visible = true;

                       Cursor.Current = Cursors.WaitCursor;

                       foreach (DataGridViewRow row in dataGridView1.Rows)
                       {
                           // select current row
                           dataGridView1.Rows[row.Index].Selected = true;
                           dataGridView1.Update();
                           currentIndex = row.Index;

                           if (row.Index.Equals(numRows))
                           {
                               // loop exceeded index range
                               break;
                           }
                           // run all scripts matching selected environment or where environment is QUAT_ALL
                           else if (proceed == true && (row.Cells[3].Value.ToString() == comboBox5.SelectedItem.ToString() || row.Cells[3].Value.ToString() == "QUAT_ALL"))
                           {
                               runScript(sender, row, lastIndexInView, e);
                               executionSummary("record");
                           }

                           if (proceed == false)
                           {   
                               // proceed will be set to false if the user chooses to edit a script mid-run
                               break;
                           }

                           // if this row is the last row of the displayed records, restart the application
                           if (row.Index == numRows)
                           {
                               toExecLog("Run sequence complete.");
                               tabControl1.Update();
                               toolStripLabel3.Text = "Execution Complete";
                               button4_Click(sender, e);
                           }
                       }
                       Cursor.Current = Cursors.Default;*/
        }



        // ------------------------------------------- Run a defined range of scripts (or resume from paused run) ----------------------------------------
        private async Task runMultiple(object sender, int numRows, int lastIndexInView, EventArgs e, int startingIndex)
        {
            toolStripLabel3.Text = "Executing scripts...";
            int result = 0;
            totalProcesses = numRows - startingIndex;
            toolStripProgressBar1.Maximum = totalProcesses + progressPosition;
            toolStripProgressBar1.Value = progressPosition;

            // updating progress value and percent is necessary after pausing the run to edit a script (UPDATE before resume will change the progress to match a successful update)
            progress = toolStripProgressBar1.Value;
            pp1 = progress / toolStripProgressBar1.Maximum;
            pp2 = (int)(pp1 * 100);
            toolStripLabel2.Text = pp2.ToString() + "%";
            toolStrip3.Update();

            toExecLog("Remaining processes: " + totalProcesses);

            // verify that the selected row's Script Order matches the Script Order entered in textBox5, if a user manually types the S.O. number into the textBox, find and select the row to match before beginning execution
            if (radioButton3.Checked == true && textBox5.Text != dataGridView1.SelectedRows[0].Cells[0].Value.ToString())
            {
                result = searchForScriptOrder(textBox5.Text); // returns the index where the S.O. is found, or returns -1 if no S.O. found, or -2 if there are duplicate S.O.'s
            }

            if (result > -1)
            {

                startingIndex = dataGridView1.SelectedRows[0].Index;

                Cursor.Current = Cursors.WaitCursor;

                // loop to run scripts within given script order range
                for (int i = startingIndex; i < numRows; i++)
                {
                    // select current row
                    dataGridView1.Rows[i].Selected = true;
                    dataGridView1.Update();

                    // set the selected row to the current index
                    dataGridView1.SelectedRows[0].Index.Equals(i);

                    // run all scripts matching the selected Run Environment or where script_environment is QUAT_ALL
                    if (proceed == true && (dataGridView1.SelectedRows[0].Cells[3].Value.ToString() == comboBox5.SelectedItem.ToString() || dataGridView1.SelectedRows[0].Cells[3].Value.ToString() == "QUAT_ALL"))
                    {
                        await Task.Run(() => runScript(sender, dataGridView1.SelectedRows[0], lastIndexInView, e));
                    }

                    if (proceed == false)
                    {
                        break;
                    }

                    updateProgressBar();
                }
                Cursor.Current = Cursors.Default;
            }
            else if (result == -1)
            {
                MessageBox.Show("No matching script order found.");
                return;
            }
            else if (result == -2)
            {
                MessageBox.Show("There is more than one record with that Script Order. Please select the record to start at.");
                return;
            }
            else
            {
                button3.Enabled = true;
                button4.Enabled = true;
                return;
            }
        }



        // --------------------------------------------------------- Run script(s) ------------------------------------------------------------
        private void runScript(object sender, DataGridViewRow currentRow, int lastIndexInView, EventArgs e)
        {
            Invoke(new Action(() =>
            {
                toExecLog("\n____________________________________________" +
                "\nRunning Script Order: " + currentRow.Cells[0].Value.ToString() + ", ID: " + currentRow.Cells[1].Value.ToString() +
                "\n------------------------------------------------------------------------------");
            }));

            Invoke(new Action(() =>
            {
                toExecLog("Record Index: " + currentRow.Index + " | " + lastIndexInView);
            }));

            Invoke(new Action(() =>
            {
                richTextBox2.Update();
            }));

            string query = currentRow.Cells[5].Value.ToString();
            try
            {
                using (conn = initQatConnection())
                {
                    try
                    {
                        using (cmd = new SqlCommand(query, conn))
                        {
                            conn.Open();
                            Invoke(new Action(() =>
                             {
                                 richTextBox2.AppendText("Connected to " + comboBox5.SelectedItem.ToString());
                             }));

                            try
                            {
                                var result = cmd.ExecuteNonQuery();
                                Convert.ToInt32(result);
                                Invoke(new Action(() =>
                                {
                                    richTextBox2.AppendText("Run Result: " + result);
                                }));
                            }
                            catch (SqlException ex)
                            {
                                logError(sender, ex, currentRow, e);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (string.IsNullOrWhiteSpace(currentRow.Cells[5].Value.ToString()))
                            MessageBox.Show(ex.Message, "ERROR: There is an empty script entry at ID:" + currentRow.Cells[1].Value.ToString());
                        else
                            MessageBox.Show(ex.Message, "From runScript:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (SqlException ex)
            {
                logError(sender, ex, currentRow, e);
            }

            Invoke(new Action(() =>
            {
                richTextBox2.Update();
            }));
        }



        // ------------------------------------------------ Update existing script --------------------------------------------------
        private void updateRecord(object sender, EventArgs e)
        {
            // progress bar update
            toolStripProgressBar1.Value = 0;
            toolStripProgressBar1.Maximum = 1;


            using (conn = initQatConnection())
            {
                string query = "UPDATE dbo.refresh_testing SET Script=@script, script_function=@script_function, script_environment=@script_environment, script_description=@script_description, script_order=@script_order WHERE ID = @id";
                using (cmd = new SqlCommand(query, conn))
                {
                    toolStripLabel3.Text = "Update in progress...";
                    cmd.Parameters.AddWithValue("id", selectedID);
                    cmd.Parameters.AddWithValue("script", richTextBox1.Text);
                    cmd.Parameters.AddWithValue("script_function", comboBox2.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("script_environment", comboBox3.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("script_description", textBox3.Text);
                    cmd.Parameters.AddWithValue("script_order", Convert.ToDecimal(textBox1.Text));

                    toExecLog("Updating record...");

                    conn.Open();
                    int result = cmd.ExecuteNonQuery();

                    // Check error
                    if (result < 0)
                    {
                        toExecLog("Could not update script. Check script for errors. Result = " + result);
                    }
                    else
                    {
                        toExecLog("Script updated successfully.");
                        toolStripLabel3.Text = "Script updated successfully.";
                        toolStripLabel3.Visible = true;
                        toolStripLabel2.Visible = true;
                        updateProgressBar();
                        tabControl1.Update();
                    }
                    toExecLog("\n------------------------------------------------------------------------------");
                }
            }
        }



        // ------------------------------------------------ Insert a new script --------------------------------------------------
        private void insertIntoTable(object sender, EventArgs e)
        {
            using (conn = initQatConnection())
            {
                string query = "INSERT INTO dbo.refresh_testing (Script, script_function, script_environment, script_description, script_order) VALUES (@Script, @script_function, @script_environment, @script_description, @script_order)";
                using (cmd = new SqlCommand(query, conn))
                {
                    toolStripLabel3.Text = "Insert in progress...";
                    cmd.Parameters.AddWithValue("@Script", richTextBox1.Text);
                    cmd.Parameters.AddWithValue("@script_function", comboBox2.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@script_environment", comboBox3.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@script_description", textBox3.Text);
                    cmd.Parameters.AddWithValue("@script_order", Convert.ToDecimal(textBox1.Text));

                    toExecLog("Inserting record...");

                    conn.Open();
                    int result = cmd.ExecuteNonQuery();

                    // Check error
                    if (result < 0)
                    {
                        toExecLog("Could not insert script. Check script for errors. Result = " + result);
                    }
                    else
                    {
                        toExecLog("Script inserted successfully.");
                        toolStripLabel3.Text = "Script inserted successfully.";
                        updateProgressBar();
                        toolStripLabel3.Visible = true;
                        tabControl1.Update();
                    }
                    toExecLog("\n------------------------------------------------------------------------------");

                    toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                    toolStripLabel2.Visible = true;
                }
            }

        }




        // END: --------------------------------------------------------- EXECUTION STATEMENTS --------------------------------------------------------------------------------------


        private void logError(object sender, SqlException exception, DataGridViewRow currentScriptRow, EventArgs e)
        {
            executionSummary("error");

            string selectedItem = "";
            if (comboBox5.InvokeRequired)
            {
                comboBox5.Invoke(new MethodInvoker(delegate { selectedItem = comboBox5.SelectedItem.ToString(); }));
            }
            else
                selectedItem = comboBox5.SelectedItem.ToString();

            string currQAT = null;

            // if selected Run Environment is not BRMQATONTDBS07, force conncection for error logging (dbo.refresh_error_log is in BRMQATONTDBS07), then return to previous selection
            if (selectedItem != "BRMQATONTDBS07")
            {
                currQAT = selectedItem;
                toExecLog("Current QAT connection: " + currQAT);
                selectedItem = "BRMQATONTDBS07";
                toExecLog("Connecting to " + selectedItem + " for error logging...");
            }


            // inserts a record detailing the error and responsible script 
            using (conn = initQatConnection())
            {
                string query = "INSERT INTO dbo.refresh_error_log (error_message, script_order, error_line, script_environment, error_time, error_procedure, script_id) VALUES (@a1, @a2, @a3, @a4, @a5, @a6, @a7)";

                // pass error  attributes to table dbo.refresh_error_log
                using (cmd = new SqlCommand(query, conn))
                {
                    DateTime date = DateTime.Now;

                    //TODO: validate the index of the error in Errors[] collection to make sure that the error concerned is always at 0 or the only element in the array 
                    cmd.Parameters.AddWithValue("@a1", exception.Errors[0].Message);
                    cmd.Parameters.AddWithValue("@a2", Convert.ToDecimal(currentScriptRow.Cells[0].Value.ToString()));
                    cmd.Parameters.AddWithValue("@a3", exception.LineNumber);
                    cmd.Parameters.AddWithValue("@a4", currentScriptRow.Cells[3].Value.ToString());
                    cmd.Parameters.AddWithValue("@a5", date);
                    cmd.Parameters.AddWithValue("@a6", exception.Errors[0].Procedure);
                    cmd.Parameters.AddWithValue("@a7", currentScriptRow.Cells[1].Value.ToString());

                    conn.Open();
                    try
                    {
                        toExecLog("Logging error...");
                        var result = cmd.ExecuteNonQuery();
                        Convert.ToInt32(result);

                        // Check error
                        if (result < 0)
                            toExecLog("Error logging failed. Result = " + result);
                        else
                            toExecLog("Error logged");
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message, "From LogError");
                        toExecLog(ex.Message);
                    }
                }
            }
            refreshErrorLog();
            if (currQAT != "BRMQATONTDBS07" && currQAT != null)
            {
                comboBox5.SelectedItem = currQAT;
                toExecLog("Connection restored to: " + comboBox5.SelectedItem.ToString());
            }


            // an Ignore All response from error dialog will prevent future error dialogs for this run session
            toExecLog("Ignore All status: " + ignoreAll);

            if (ignoreAll == false)
            {
                DataGridView dgv = dataGridView2 as DataGridView;
                if (dgv != null)
                {
                    selectLastErrorLogged();

                    // object to pass selected row
                    DataGridViewRow errorRow = dgv.SelectedRows[0];

                    // open error dialog form and store user's response
                    string response = openErrorPrompt(errorRow);

                    // handling user's response from error dialog 
                    if (response == "ignore")
                        toExecLog("Ignoring error");
                    else if (response == "ignore all")
                    {
                        ignoreAll = true;
                        toExecLog("Ignoring all errors for this run session\n\t -errors will still be logged");
                    }
                    else if (response == "edit script")
                    {
                        toExecLog("Entering edit mode...");
                        if (radioButton1.Checked == true || radioButton3.Checked == true)
                            prepareForUpdateAndResume(sender, e, exception, currentScriptRow);

                        else if (radioButton4.Checked == true)
                        {
                            richTextBox1.Enabled = true;
                            tabControl1.SelectedIndex = 0;
                            highlightLine(richTextBox1, Convert.ToInt32(errorRow.Cells[3].Value), Yellow);
                            proceed = true;
                        }

                    }
                    else if (response == "quit")
                    {
                        button4_Click(sender, e);
                        proceed = false;
                    }
                    else if (response == "failure")
                    {
                        MessageBox.Show(response);
                    }
                    else
                    {
                        response = "quit";
                        toExecLog(response + "\nAborting...");
                        proceed = false;
                    }
                }
                else
                    MessageBox.Show("The error log is empty");
            }
            Console.WriteLine("Ignore all is true");
        }



        private int searchForScriptOrder(string scriptOrder)
        {
            int index = 0;
            int matchCount = 0;

            dataGridView1.Rows[index].Selected = true;

            // loop to search for mathcing Script Order
            while (index < dataGridView1.Rows.Count)
            {
                if (dataGridView1.Rows[index].Cells[0].Value.ToString() == scriptOrder)
                {
                    // select the row where the Script Order matches the textBox5 Script Order
                    dataGridView1.Rows[index].Selected = true;
                    matchCount++;
                }
                index++;
            }
            if (matchCount == 0)
            {
                return -1;
            }
            else if (matchCount > 1)
            {
                return -2;
            }
            else
            {
                return dataGridView1.SelectedRows[0].Index;
            }
        }



        private void prepareForUpdateAndResume(object sender, EventArgs e, SqlException exception, DataGridViewRow currentScriptRow)
        {
            // store value of progress bar 
            progressPosition = toolStripProgressBar1.Value;

            // switch to tab where script can be modified
            tabControl1.SelectedIndex = 0;

            // prepare script in richTextBox 
            richTextBox1.Enabled = true;
            richTextBox1.Text = currentScriptRow.Cells[5].Value.ToString();
            colorizeScript();
            highlightLine(richTextBox1, exception.LineNumber, Yellow);

            // modify the Action comboBox to only contain "UPDATE" and "UPDATE & RESUME"
            if (comboBox1.Items.Contains("INSERT INTO"))
                comboBox1.Items.Remove("INSERT INTO");
            if (comboBox1.Items.Contains("RUN"))
                comboBox1.Items.Remove("RUN");
            if (!comboBox1.Items.Contains("UPDATE & RESUME"))
                comboBox1.Items.Add("UPDATE & RESUME");
            comboBox1.SelectedItem = "UPDATE & RESUME";

            // stop the RUN sequence
            proceed = false;

            // set placeholder to index of record being fixed
            placeholder = Convert.ToInt32(dataGridView1.SelectedRows[0].Index);
            toExecLog("Execution stopped at row index: " + currentScriptRow.Index);
            textBox5.Text = currentScriptRow.Index.ToString();
            radioButton3.Checked = true;
            setSelectedValues(sender);
        }



        private void selectLastErrorLogged()
        {
            // select most recent (last) error record         
            int lastRow = dataGridView2.Rows.Count - 1;
            dataGridView2.Rows[lastRow].Selected = true;
        }



        private void highlightLine(RichTextBox rtb1, int index, Color color)
        {
            index -= 1;
            var lines = rtb1.Lines;
            if (index < 0 || index >= lines.Length)
                return;
            var start = rtb1.GetFirstCharIndexFromLine(index);
            var length = lines[index].Length;
            toExecLog("Highlight line: " + index + "\nStart :" + start + "\nLength: " + length);
            rtb1.SelectionBackColor = color;
            rtb1.Select(start, length);
            rtb1.SelectionBackColor = color;
        }



        private string openErrorPrompt(DataGridViewRow errorRowSelected)
        {
            DataGridView dgv = dataGridView2 as DataGridView;
            if (dgv != null && dgv.SelectedRows.Count > 0)
            {
                // set form values equal to values from selected row
                toExecLog("Opening error dialog...");
                ErrorPrompt f2 = new ErrorPrompt();
                f2.colorScheme = lightSwitch;
                f2.errorMessage = errorRowSelected.Cells[4].Value.ToString();
                f2.errorLine = Convert.ToInt32(errorRowSelected.Cells[3].Value);
                f2.procedure = errorRowSelected.Cells[5].Value.ToString();
                f2.scriptOrder = errorRowSelected.Cells[0].Value.ToString();
                f2.id = errorRowSelected.Cells[1].Value.ToString();
                f2.env = errorRowSelected.Cells[2].Value.ToString();

                // show the form and return the user's error-handling choice
                f2.ShowDialog();
                string errorOption = f2.option;
                toExecLog("Your response: " + errorOption);
                f2.Close();
                return errorOption;
            }
            string failure = "Failed to open error dialog: selected error record is null or nothing was selected";
            return failure;
        }



        private void execute_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // call execution if ENTER key is pressed
                button3_Click(sender, e);
            }
        }



        private void searchTable(KeyEventArgs e)
        {
            dt = new DataTable("refresh_testing");
            DataView dv = dt.DefaultView;
            try
            {
                using (conn = initQatConnection())
                {
                    conn.Open();
                    using (SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM refresh_testing ORDER BY script_order ASC", conn))
                    {
                        toExecLog("Loading table...");
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;

                        // search bar feature, searches selected column for toolStriptextBox1 text upon ENTER key_up and BACKSPACE key_up
                        //NOTE: other Key Events (ex. key_down) combined with this particular search config will cause search to lag by one keystroke

                        if (toolStripComboBox1.SelectedIndex == 0)
                        {
                            dv.RowFilter = string.Format("CONVERT(script_order, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }

                        else if (toolStripComboBox1.SelectedIndex == 1)
                        {
                            dv.RowFilter = string.Format("CONVERT(ID, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }

                        else if (toolStripComboBox1.SelectedIndex == 2)
                        {
                            dv.RowFilter = string.Format("CONVERT(script_function, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }

                        else if (toolStripComboBox1.SelectedIndex == 3)
                        {
                            dv.RowFilter = string.Format("CONVERT(script_environment, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }

                        else if (toolStripComboBox1.SelectedIndex == 4)
                        {
                            dv.RowFilter = string.Format("CONVERT(script_description, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }

                        else if (toolStripComboBox1.SelectedIndex == 5)
                        {
                            dv.RowFilter = string.Format("CONVERT(Script, 'System.String') LIKE '%{0}%'", toolStripTextBox1.Text.Trim());
                            dataGridView1.DataSource = dv.ToTable();
                        }
                    }
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void toolStripTextBox1_KeyUp(object sender, KeyEventArgs e)
        {
            // search table up pressing ENTER key
            if (e.KeyCode == Keys.Enter)
            {
                searchTable(e);
            }
            // table will search upon a BACKSPACE
            else if (e.KeyCode == Keys.Back)
            {
                searchTable(e);
            }
        }



        private void colorizeScript()
        {
            // SQL text formatting
            toolStripLabel3.Text = "Formatting script...";
            toolStripLabel3.Visible = true;
            Manoli.Utils.CSharpFormat.TsqlFormat clrz = new Manoli.Utils.CSharpFormat.TsqlFormat();
            richTextBox1.Rtf = clrz.FormatString(richTextBox1.Text);
            toolStripLabel3.Visible = false;
        }



        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // allows for digits, decimal, and backspace / delete entries
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != 46)
            {
                e.Handled = true;
                return;
            }

            // allows for only one decimal to be entered
            if (e.KeyChar == 46)
            {
                if ((sender as TextBox).Text.IndexOf(e.KeyChar) != -1)
                    e.Handled = true;
            }

            // allows only two digits after decimal
            if (Regex.IsMatch(textBox1.Text, @"\.\d\d") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }



        private void decryptAccess()
        {
            // calls decrypt function in DB "master" to check user's permission to run scripts on user selected QAT
            // decrypt function checks to see if system user name is found in table for corresponding QAT connection. If matching record exists, 1 is returned. 
            object result = 0;
            Cursor.Current = Cursors.WaitCursor;

            try
            {
                if (conn != null && conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    toExecLog("Connection was open, now closed");
                }

                using (conn = initQatConnection())
                {
                    string query = @"DECLARE @hasAccess INT = 0, @sql NVARCHAR(MAX) BEGIN TRY set @sql = 'SELECT @hasAccess = [master].[dbo].[fn_get_user_access] (SUSER_SNAME() , ''OnTrakProd'')' EXEC sp_executeSQL @sql, N'@hasAccess int output', @hasAccess output SELECT @hasAccess END TRY BEGIN CATCH set @sql = 'SELECT @hasAccess = [master].[dbo].[fn_get_user_access] ( SUSER_SNAME() , ''OnTrakProd'', NULL)' EXEC sp_executeSQL @sql, N'@hasAccess int output', @hasAccess output SELECT @hasAccess END CATCH";
                    SqlCommand cmd = new SqlCommand(query, conn);
                    conn.Open();
                    cmd = new SqlCommand(query, conn);
                    result = cmd.ExecuteScalar();

                    // Decrypt Access response
                    if ((int)result != 1)
                    {
                        toolStripLabel4.Visible = true;
                        toolStripLabel3.Text = "You do not have DeCrypt access in database " + comboBox5.SelectedItem.ToString();
                        toolStripLabel3.Visible = true;
                        toolStripLabel3.IsLink = true;
                    }
                    else
                    {
                        toolStripLabel3.Visible = false;
                        toolStripLabel4.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error openning connection: ");
            }
            Cursor.Current = Cursors.Default;
        }



        private SqlConnection initQatConnection()
        {
            // this function initializes and then returns a new SqlConnection
            // if the action selected is RUN, this function will establish a connection based on the User's Run Environment selection 
            // otherwise it will connect to QAT7 for all other purposes

            int index = 0;
            if (comboBox5.InvokeRequired)
            {
                // comboBox5.Invoke(new Action(() => comboBox5.SelectedIndex = index));
                comboBox5.Invoke(new MethodInvoker(delegate { index = comboBox5.SelectedIndex; }));
            }
            else
                index = comboBox5.SelectedIndex;
            
            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();

            // RM connecion strings 

            // if RUN is selected along with a Run Environment from combobox5, connect to the selected Run Environment
            if (actionSelected == "RUN" && index != -1)
            {
                // no selection
                if (index == -1)
                    return null;
                // brmqatontdbs02 RM
                else if (index == 0)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs02;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs03 RM
                else if (index == 1)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs03;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs04 RM
                else if (index == 2)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs04;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs05 RM
                else if (index == 3)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs05;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs06 RM
                else if (index == 4)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs06;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs07 RM
                else if (index == 5)
                {
                    conn = new SqlConnection(@"Data Source=Brmqatontdbs07;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs11 RM
                else if (index == 6)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs11;Initial Catalog=RM;Integrated Security=True");
                }
                else
                {
                    toExecLog("The selected QAT Environment " + comboBox5.GetItemText(comboBox5.SelectedItem) + " does not have a defined connection.");
                    return null;
                }
            }
            else
            {
                // brmqatontdbs07 RM
                conn = new SqlConnection(@"Data Source=Brmqatontdbs07;Initial Catalog=RM;Integrated Security=True");
            }
            return conn;
        }



        public void loadData()
        {
            if (conn != null && conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

            // connect and fill table with dbo.refresh_testing data
            dt = new DataTable("refresh_testing");
            try
            {
                using (conn = initQatConnection())
                {
                    conn.Open();

                    using (SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM refresh_testing ORDER BY script_order ASC", conn))
                    {
                        toExecLog("Loading table...");
                        da.Fill(dt);
                        dataGridView1.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error establishing connection to table", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            refreshErrorLog();
        }




        private void setSelectedValues(object sender)
        {
            // this view is used for setting UPDATE & RUN fields to the values of the selected row
            DataGridView dgv = dataGridView1 as DataGridView;

            if ((dgv != null && dgv.SelectedRows.Count > 0) && (dgv.SelectedRows[0].Selected = true) && actionSelected != "INSERT INTO")
            {
                DataGridViewRow row = dgv.SelectedRows[0];

                if (actionSelected == "UPDATE" || (actionSelected == "RUN" && radioButton4.Checked == true) || actionSelected == "UPDATE & RESUME")
                {
                    // set selected Script Order number to textBox
                    textBox1.Text = row.Cells[0].Value.ToString();


                    // set selected ID number to textBox, this disabled and cannot be changed
                    selectedID = row.Cells[1].Value.ToString();
                    textBox2.Text = selectedID;


                    // set selected row value to Function comboBox
                    for (int i = 0; i < comboBox2.Items.Count; i++)
                    {
                        if (row.Cells[2].Value.ToString() == comboBox2.GetItemText(comboBox2.Items[i]))
                        {
                            comboBox2.SelectedIndex = i;
                            break;
                        }
                    }

                    // set selected row value to Script Environment comboBox
                    for (int i = 0; i < comboBox3.Items.Count; i++)
                    {
                        if (row.Cells[3].Value.ToString() == comboBox3.GetItemText(comboBox3.Items[i]))
                        {
                            comboBox3.SelectedIndex = i;
                            break;
                        }
                    }

                    // set selected Description to textBox
                    textBox3.Text = row.Cells[4].Value.ToString();

                    // set selected Script query to Script View tab
                    richTextBox1.Text = row.Cells[5].Value.ToString();
                    colorizeScript();
                }

                if (radioButton4.Checked == true)
                {
                    // set Run Environment equal to the selected Script Environment if it is not QUAT_ALL or empty
                    for (int i = 0; i < comboBox5.Items.Count; i++)
                    {
                        if (row.Cells[3].Value.ToString() == comboBox5.GetItemText(comboBox5.Items[i]))
                        {
                            comboBox5.SelectedIndex = i;
                            break;
                        }
                        else
                            comboBox5.SelectedIndex = -1;
                    }
                }

                if (radioButton3.Checked == true)
                {
                    textBox5.Text = row.Cells[0].Value.ToString();
                }
            }
        }



        private void button7_Click(object sender, EventArgs e)
        {
            // CLEAR ERROR LOG button
            using (conn = initQatConnection())
            {
                string query = "DELETE FROM dbo.refresh_error_log";

                using (cmd = new SqlCommand(query, conn))
                {
                    conn.Open();
                    int result = cmd.ExecuteNonQuery();

                    // Check error
                    if (result < 0)
                        toExecLog("Deletion failed.");
                    else
                    {
                        toolStripLabel4.Visible = false;
                        toolStripLabel3.Text = "Error Log cleared";
                        toolStripLabel3.Visible = true;
                        toExecLog("Error log cleared");
                    }
                    loadData();
                }
            }
        }



        private void refreshErrorLog()
        {
            if (conn != null && conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

            // connect and fill table with dbo.refresh data
            dt = new DataTable("refresh_error_log");
            try
            {
                using (conn = initQatConnection())
                {
                    conn.Open();

                    using (SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM refresh_error_log ORDER BY error_time ASC", conn))
                    {
                        toExecLog("Loading error log...");
                        da.Fill(dt);
                        dataGridView2.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error establishing connection to table", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            toExecLog("Load complete" +
            "\n------------------------------------------------------------------------------");
        }



        private void button8_Click(object sender, EventArgs e)
        {
            // INSPECT ERROR button

            // creates an object for selected row in error log
            DataGridView dgv = dataGridView2 as DataGridView;

            try
            {
                DataGridViewRow row = dgv.SelectedRows[0];
                if (row.Cells[0].Value != null)
                {
                    // set form values equal to values from selected row
                    toExecLog("Inspecting error...");
                    InspectError form = new InspectError();
                    form.colorScheme = lightSwitch;
                    form.errorMessage = row.Cells[4].Value.ToString();
                    form.errorLine = Convert.ToInt32(row.Cells[3].Value);
                    form.procedure = row.Cells[5].Value.ToString();
                    form.scriptOrder = row.Cells[0].Value.ToString();
                    form.id = row.Cells[1].Value.ToString();
                    form.env = row.Cells[2].Value.ToString();

                    // show the form and return the user's error-handling choice
                    form.ShowDialog();
                    string response = form.option;
                    form.Close();
                    if (response == "edit script")
                    {
                        editScript(sender, e, row);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Please select a valid error record");
            }
        }



        private void editScript(object sender, EventArgs e, DataGridViewRow row)
        {
            int scriptID = Convert.ToInt32(row.Cells[1].Value);
            tabControl1.SelectedIndex = 0;
            richTextBox1.Enabled = true;

            // search the script table for matching script
            DataGridView grid = dataGridView1 as DataGridView;
            for (int i = 0; i < grid.Rows.Count; i++)
            {
                if (scriptID == Convert.ToInt32(grid.Rows[i].Cells[1].Value))
                {
                    richTextBox1.Text = grid.Rows[i].Cells[5].Value.ToString();
                    colorizeScript();
                    highlightLine(richTextBox1, Convert.ToInt32(row.Cells[3].Value), Yellow);
                    break;
                }
            }
        }


        public void toExecLog(string input)
        {
            if (richTextBox2.InvokeRequired)
            {
                richTextBox2.Invoke(new MethodInvoker(delegate { richTextBox2.AppendText("\r" + input); }));
            }
            else
            {
                // outputs to Execution Log rtb
                richTextBox2.AppendText("\r" + input);
                richTextBox2.ScrollToCaret();
            }
        }


        private bool fieldCheck(object sender, EventArgs e)
        {
            if (actionSelected == "INSERT INTO" || actionSelected == "UPDATE")
            {
                if (string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    MessageBox.Show("A script order is required (int)");
                    return false;
                }

                else if (comboBox2.SelectedIndex == -1)
                {
                    MessageBox.Show("A function must be selected.");
                    return false;
                }

                else if (comboBox3.SelectedIndex == -1)
                {
                    MessageBox.Show("A script environment must be selected.");
                    return false;
                }

                else if (string.IsNullOrWhiteSpace(textBox3.Text))
                {
                    MessageBox.Show("A description is required.");
                    return false;
                }
                else if (string.IsNullOrWhiteSpace(richTextBox1.Text))
                {
                    DialogResult dr = MessageBox.Show("A script must be loaded. Would you like to use the selected script?", "ATTENTION", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    if (dr == DialogResult.Yes)
                    {
                        if (dataGridView1 != null && dataGridView1.SelectedRows.Count > 0)
                        {
                            tabControl1.SelectedIndex = 0;
                            // set selected Script query to preview tab
                            richTextBox1.Text = dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
                            return false;
                        }
                    }
                    else
                        return false;
                }
            }
            else if (actionSelected == "RUN")
            {
                if ((radioButton1.Checked == false && radioButton3.Checked == false && radioButton4.Checked == false))
                {
                    MessageBox.Show("Please select a RUN option.");
                    return false;
                }
                else if (comboBox5.SelectedIndex == -1)
                {
                    MessageBox.Show("A run environment must be selected.");
                    return false;
                }
                else if (string.IsNullOrWhiteSpace(richTextBox1.Text) && radioButton4.Checked == true)
                {
                    // prompt user if selected script field is blank upon execution, offer to parse from selection
                    DialogResult diRes = MessageBox.Show("A script must be loaded. Would you like to use the selected script?", "ATTENTION", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    if (diRes == DialogResult.Yes)
                    {
                        DataGridView dgv = sender as DataGridView;
                        if (dgv != null && dgv.SelectedRows.Count > 0)
                        {
                            // set selected Script query to view tab
                            toolStripButton3_Click(sender, e);
                            return false;
                        }
                        else
                        {
                            dataGridView1.Rows[0].Selected = true;
                            toolStripButton3_Click(sender, e);
                            return false;
                        }
                    }
                    else
                        return false;
                }
            }
            else
            {
                if (comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("An action must be selected.");
                    return false;

                }
            }
            return true;
        }



        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                clearFields(sender, e);
                comboBox5.Enabled = true;
                richTextBox1.Text = "-- ALL SCRIPTS WITH ENVIRONMENT [Select Environment] OR QUAT_ALL WILL BE RAN";
                toggleFieldVisibility(0);
            }
            else
            {
                toolStripButton3_Click(sender, e);
            }
        }



        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                clearFields(sender, e);
                textBox5.Enabled = true;
                if (!lightSwitch)
                    textBox5.BackColor = Color.Lavender;
                else
                    textBox5.BackColor = Color.White;
                toggleFieldVisibility(0);
                comboBox5.Enabled = true;

                try
                {
                    if (dataGridView1.SelectedRows[0].Selected == false)
                    {
                        dataGridView1.Rows[0].Selected = true;
                        setSelectedValues(sender);
                    }
                    else
                        textBox5.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                }
                catch
                {
                    dataGridView1.Rows[0].Selected = true;
                    setSelectedValues(sender);
                }
            }
            else
            {
                textBox5.Enabled = false;
                textBox5.Text = null;
                if (!lightSwitch)
                    textBox5.BackColor = Color.DimGray;
                else
                    textBox5.BackColor = Color.LightGray;
            }
        }



        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (actionSelected == "RUN" && comboBox5.SelectedIndex != -1)
            {
                decryptAccess();
                if (radioButton1.Checked == true)
                    richTextBox1.Text = "-- ALL SCRIPTS WITH ENVIRONMENT " + comboBox5.SelectedItem.ToString() + " OR QUAT_ALL WILL BE RAN";
            }
        }



        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 1)
                label16.Text = "RM.dbo.refresh_error_log      (BRMQATONTDBS07)";
            else
                label16.Text = "RM.dbo.refresh_testing      (BRMQATONTDBS07)";
        }



        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            comboBox5.Enabled = true;

            if (radioButton4.Checked == true)
            {
                toggleFieldVisibility(1);
                setSelectedValues(sender);
                comboBox5.Visible = true;
                label11.Enabled = true;
                label5.Enabled = true;
                label4.Enabled = true;
                label10.Enabled = true;
                label14.Enabled = true;
                label13.Enabled = true;
                textBox1.ReadOnly = true;
                textBox3.ReadOnly = true;
                textBox1.Enabled = true;
                textBox3.Enabled = true;
            }
            else
            {
                textBox1.ReadOnly = false;
                textBox3.ReadOnly = false;
                textBox1.Enabled = false;
                textBox3.Enabled = false;
            }
        }



        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            setSelectedValues(sender);
        }



        private void clearFields(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;

            if (richTextBox1.Enabled == true)
                richTextBox1.Enabled = false;
        }



        private void toggleFieldVisibility(int x)
        {
            // RUN settings
            if (x == 0)
            {
                textBox1.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                label10.Visible = false;
                label11.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                comboBox2.Visible = false;
                comboBox3.Visible = false;
                comboBox5.Visible = true;
            }
            // INSERT and UPDATE settings
            else if (x == 1)
            {
                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                label10.Visible = true;
                label11.Visible = true;
                label13.Visible = true;
                label14.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                comboBox2.Visible = true;
                comboBox3.Visible = true;
                comboBox5.Visible = true;
            }
        }



        private void flipLightSwitch(bool on)
        {
            // color scheme settings
            // off (dark) by default
            if (on)
            {
                // light color scheme
                textBox1.BackColor = Color.White;
                textBox2.BackColor = Color.White;
                textBox3.BackColor = Color.White;
                textBox1.ForeColor = Color.Black;
                textBox2.ForeColor = Color.Black;
                textBox3.ForeColor = Color.Black;
                if (radioButton3.Checked)
                    textBox5.BackColor = Color.White;
                else
                    textBox5.BackColor = Color.LightGray;
                label1.ForeColor = Color.Black;
                label2.ForeColor = Color.Black;
                toolStripLabel3.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label16.ForeColor = Color.Black;
                label17.ForeColor = Color.Black;
                toolStripLabel2.ForeColor = Color.Black;
                label10.ForeColor = Color.Black;
                label11.ForeColor = Color.Black;
                label12.ForeColor = Color.Black;
                label13.ForeColor = Color.Black;
                label14.ForeColor = Color.Black;
                label4.ForeColor = Color.Black;
                label5.ForeColor = Color.Black;
                comboBox1.BackColor = Color.Orange;
                comboBox2.BackColor = Color.Orange;
                comboBox3.BackColor = Color.Orange;
                toolStripComboBox1.BackColor = Color.Orange;
                toolStripTextBox1.BackColor = Color.White;
                comboBox5.BackColor = Color.Orange;
                toolStrip2.BackColor = Color.LightGray;
                toolStripLabel1.ForeColor = Color.Black;
                this.BackColor = Color.LightGray;
                dataGridView1.DefaultCellStyle.BackColor = Color.White;
                dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DarkOrange;
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Moccasin;
                dataGridView1.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.DarkOrange;
                dataGridView2.DefaultCellStyle.BackColor = Color.White;
                dataGridView2.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.DarkOrange;
                dataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.DarkOrange;
                dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Orange;
                dataGridView2.BackgroundColor = Color.White;
                dataGridView1.BackgroundColor = Color.White;
                button7.BackColor = Color.Moccasin;
                button7.ForeColor = Color.Black;
                button8.BackColor = Color.Moccasin;
                button8.ForeColor = Color.Black;
                toolStripButton4.BackColor = Color.Orange;
                toolStripButton4.ForeColor = Color.Black;
                toolStripButton3.BackColor = Color.Moccasin;
                toolStripButton3.ForeColor = Color.Black;
                toolStripButton2.BackColor = Color.Moccasin;
                toolStripButton2.ForeColor = Color.Black;
                button3.BackColor = Color.ForestGreen;
                button4.BackColor = Color.Gold;
                button5.BackColor = Color.Firebrick;
                this.ForeColor = Color.Black;
                gradientPanel1.ColorLeft = Color.White;
                gradientPanel1.ColorRight = Color.DarkOrange;
                gradientPanel2.ColorLeft = Color.DarkOrange;
                gradientPanel2.ColorRight = Color.White;
                richTextBox2.BackColor = Color.White;
                richTextBox2.ForeColor = Color.Black;
                radioButton1.ForeColor = Color.Black;
                radioButton3.ForeColor = Color.Black;
                radioButton4.ForeColor = Color.Black;
                toolStrip3.BackgroundImage = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\TS3OrangeGradient.bmp");
                toolStrip2.BackgroundImage = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\TS2OrangeGradient.bmp");
            }
            else
            {
                // dark color scheme
                textBox1.BackColor = Color.DimGray;
                textBox2.BackColor = Color.DimGray;
                textBox3.BackColor = Color.DimGray;
                textBox1.ForeColor = Color.White;
                textBox2.ForeColor = Color.White;
                textBox3.ForeColor = Color.White;
                if (radioButton3.Checked)
                    textBox5.BackColor = Color.Lavender;
                else
                    textBox5.BackColor = Color.DimGray;
                label1.ForeColor = Color.Gainsboro;
                label2.ForeColor = Color.Gainsboro;
                toolStripLabel3.ForeColor = Color.Black;
                label9.ForeColor = Color.Gainsboro;
                label16.ForeColor = Color.Gainsboro;
                label17.ForeColor = Color.Gainsboro;
                toolStripLabel2.ForeColor = Color.Black;
                label10.ForeColor = Color.Gainsboro;
                label11.ForeColor = Color.Gainsboro;
                label12.ForeColor = Color.Gainsboro;
                label13.ForeColor = Color.Gainsboro;
                label14.ForeColor = Color.Gainsboro;
                label4.ForeColor = Color.Gainsboro;
                label5.ForeColor = Color.Gainsboro;
                comboBox1.BackColor = Color.PaleTurquoise;
                comboBox2.BackColor = Color.PaleTurquoise;
                comboBox3.BackColor = Color.PaleTurquoise;
                toolStripComboBox1.BackColor = Color.PaleTurquoise;
                toolStripTextBox1.BackColor = Color.Lavender;
                comboBox5.BackColor = Color.PaleTurquoise;
                toolStrip2.BackColor = Color.FromArgb(64, 64, 64);
                toolStripLabel1.ForeColor = Color.Gainsboro;
                this.BackColor = Color.FromArgb(64, 64, 64);
                dataGridView1.DefaultCellStyle.BackColor = Color.White;
                dataGridView1.DefaultCellStyle.SelectionBackColor = Color.DeepSkyBlue;
                dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
                dataGridView1.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.DeepSkyBlue;
                dataGridView2.DefaultCellStyle.BackColor = Color.MediumPurple;
                dataGridView2.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.Indigo;
                dataGridView2.RowsDefaultCellStyle.SelectionBackColor = Color.Indigo;
                dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Indigo;
                dataGridView2.BackgroundColor = Color.DimGray;
                dataGridView1.BackgroundColor = Color.DimGray;
                button7.BackColor = Color.MediumPurple;
                button7.ForeColor = SystemColors.ButtonHighlight;
                button8.BackColor = Color.MediumPurple;
                button8.ForeColor = SystemColors.ButtonHighlight;
                toolStripButton4.BackColor = Color.RoyalBlue;
                toolStripButton4.ForeColor = Color.White;
                toolStripButton3.BackColor = Color.MediumTurquoise;
                toolStripButton3.ForeColor = Color.Black;
                toolStripButton2.BackColor = Color.LimeGreen;
                toolStripButton2.ForeColor = Color.White;
                button3.BackColor = Color.LimeGreen;
                button4.BackColor = Color.MediumTurquoise;
                button5.BackColor = Color.RoyalBlue;
                this.ForeColor = Color.Black;
                gradientPanel1.ColorLeft = Color.Cyan;
                gradientPanel1.ColorRight = Color.Indigo;
                gradientPanel2.ColorLeft = Color.Indigo;
                gradientPanel2.ColorRight = Color.Cyan;
                richTextBox2.BackColor = Color.DimGray;
                richTextBox2.ForeColor = Color.Ivory;
                radioButton1.ForeColor = Color.White;
                radioButton3.ForeColor = Color.White;
                radioButton4.ForeColor = Color.White;
                toolStrip3.BackgroundImage = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\TS3gradient.bmp");
                toolStrip2.BackgroundImage = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\TS2gradient.bmp");
            }
        }



        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // off by default
            if (lightSwitch)
                lightSwitch = false;
            else
                lightSwitch = true;
            flipLightSwitch(lightSwitch);
        }



        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            // allows for digits, decimal, and backspace / delete entries
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != 46)
            {
                e.Handled = true;
                return;
            }

            // allows for only one decimal to be entered
            if (e.KeyChar == 46)
            {
                if ((sender as TextBox).Text.IndexOf(e.KeyChar) != -1)
                    e.Handled = true;
            }

            // allows only two digits after decimal
            if (Regex.IsMatch(textBox5.Text, @"\.\d\d") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }



        private void textBox5_MouseClick(object sender, MouseEventArgs e)
        {
            textBox5.ReadOnly = false;
            textBox5.Text = "";
        }



        private void textBox5_Leave(object sender, EventArgs e)
        {
            // manage textbox input to ensure that decimal values to the hundredth's place are always used

            if (string.IsNullOrWhiteSpace(textBox5.Text))
                textBox5.Text = "0.00";

            decimal so = Convert.ToDecimal(textBox5.Text);
            int decimalPlaces = BitConverter.GetBytes(decimal.GetBits(so)[3])[2];

            if (decimalPlaces == 0)
            {
                textBox5.Text = so + ".00";
            }
            else if (decimalPlaces == 1)
            {
                textBox5.Text = so + "0";
            }
        }



        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            // BROWSE button opens a file dialog to upload files
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "SQL|*.sql";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                if ((myStream = ofd.OpenFile()) != null)
                {
                    string strFileName = ofd.FileName;
                    string fileText = File.ReadAllText(strFileName);
                    richTextBox1.Text = fileText;
                }
                else
                {
                    MessageBox.Show("ERROR: could not open the file");
                }
                label1.Text = ofd.FileName;
                label1.Visible = true;
            }
        }



        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            // LOCK / UNLOCK BUTTON 
            if (richTextBox1.Enabled == false)
            {
                richTextBox1.Enabled = true;
                toolStripButton2.Image = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\icons\unlock.png");
                toolStripButton2.ToolTipText = "Lock editor";
            }
            else
            {
                richTextBox1.Enabled = false;
                toolStripButton2.Image = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\icons\lock.png");
                toolStripButton2.ToolTipText = "Unlock editor";
            }
        }



        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            // PARSE BUTTON
            DataGridView dgv = dataGridView1 as DataGridView;

            if (dgv != null && dgv.SelectedRows.Count > 0)
            {
                DataGridViewRow row = dgv.SelectedRows[0];
                if (row != null)
                {
                    richTextBox1.Text = row.Cells[5].Value.ToString();
                    colorizeScript();
                }
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            // hard code to save gradient panel as an image which can then be set as a toolStrip background image
            // oddly enough, this seems to be the easiest way wihout PhotoShop; Paint is useless.
            // button is hidden by default
            // panels are hiding behind toolStrips

            int pWidth = gradientPanel2.Size.Width;
            int pHeight = gradientPanel2.Size.Height;
            Bitmap bm = new Bitmap(pWidth, pHeight);
            gradientPanel2.DrawToBitmap(bm, new Rectangle(0, 0, pWidth, pHeight));
            bm.Save(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\TS3OrangeGradient.bmp", System.Drawing.Imaging.ImageFormat.Bmp);
        }



        private void textBox1_Leave(object sender, EventArgs e)
        {
            // manage textbox input to ensure that decimal values to the hundredth's place are always used

            if (string.IsNullOrWhiteSpace(textBox1.Text))
                textBox1.Text = "0.00";

            decimal so = Convert.ToDecimal(textBox1.Text);
            int decimalPlaces = BitConverter.GetBytes(decimal.GetBits(so)[3])[2];

            if (decimalPlaces == 0)
            {
                textBox1.Text = so + ".00";
            }
            else if (decimalPlaces == 1)
            {
                textBox1.Text = so + "0";
            }
        }



        private void richTextBox1_EnabledChanged(object sender, EventArgs e)
        {
            // change LOCK / UNLOCK button icon accordingly 
            if (richTextBox1.Enabled == true)
            {
                toolStripButton2.Image = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\icons\unlock.png");
                toolStripButton2.ToolTipText = "Lock editor";
            }
            else
            {
                toolStripButton2.Image = Image.FromFile(@"\\BRMSPNVDIFSV01\TFD_Profile\agemperline\Documents\My Pictures\icons\lock.png");
                toolStripButton2.ToolTipText = "Unlock editor";
            }
        }



        private void updateProgressBar()
        {
            if (toolStripProgressBar1.Value < toolStripProgressBar1.Maximum)
                toolStripProgressBar1.Value++;

            progress = toolStripProgressBar1.Value;
            pp1 = progress / toolStripProgressBar1.Maximum;
            pp2 = (int)(pp1 * 100);
            toolStripLabel2.Text = pp2.ToString() + "%";

            if (toolStripProgressBar1.Value == toolStripProgressBar1.Maximum)
            {
                toolStripLabel3.Visible = true;
                toolStripLabel2.Text = "100%";
                toolStripLabel3.Text = "Process Complete.";
            }
            toolStrip3.Update();
        }



        private void delay(int milliseconds)
        {
            int i = 0;
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Start();
            timer.Interval = milliseconds;
            timer.Elapsed += (s, args) => i = 1;
            while (i == 0) { };
        }



        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            // ESM and DeCrypt infolink

            ToolStripLabel toolStripLabel3 = (ToolStripLabel)sender;

            // start Internet Explorer and navigate to the URL in the tag property
            System.Diagnostics.Process.Start("IEXPLORE.EXE", toolStripLabel3.Tag.ToString());

            // set the LinkVisited propert to true to change the color
            toolStripLabel3.LinkVisited = true;

        }



        private void executionSummary(string varType)
        {
            // tally variables
            int numRecords = 0;
            int numErrors = 0;
            int errorsResolved = 0;
            string env = "";

            if (varType == "record")
                numRecords++;
            else if (varType == "error")
                numErrors++;
            else if (varType == "errorResolved")
                errorsResolved++;
            else if (varType == "display")
            {
                if (actionSelected == "INSERT INTO" || actionSelected == "UPDATE")
                {
                    env = "brmqatontdbs07";
                }
                else if (actionSelected == "RUN")
                {
                    env = comboBox5.GetItemText(comboBox5.SelectedItem).ToLower();
                }

                // print summary to Execution Log
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.SelectionColor = Color.Ivory;

                richTextBox2.AppendText("\n\n============================================");
                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.AppendText("\nEXECUTION SUMMMARY");
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText("\n----------------------------------------------------------------------------");

                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.AppendText("\nAction: ");
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Regular);
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText(actionSelected.ToLower());

                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.AppendText("\nEnvironment: ");
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Regular);
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText(env.ToLower());

                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.AppendText("\nRecords: ");
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Regular);
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText(numRecords.ToString());

                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.AppendText("\nErrors: ");
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Regular);
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText(numErrors.ToString());

                richTextBox2.SelectionColor = Color.MediumTurquoise;
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Bold);
                richTextBox2.AppendText("\nErrors Resolved: ");
                richTextBox2.SelectionFont = new Font("Arial", 9, FontStyle.Regular);
                richTextBox2.SelectionColor = Color.Ivory;
                richTextBox2.AppendText(errorsResolved.ToString());

                richTextBox2.AppendText("\n\n============================================");
            }
        }



        // START: ---------------------------------------------- Menu and contextMenuStrip Items' events ----------------------------------------------


        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            // right-click on Script table opens contextMenuStrip
            toolStripComboBox2.SelectedIndex = -1;
            toolStripComboBox3.SelectedIndex = -1;
            toolStripComboBox4.SelectedIndex = -1;

            if (e.RowIndex != -1 && e.Button == MouseButtons.Right)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[1];
                contextMenuStrip1.Show(Cursor.Position);
            }
        }

        private void deleteRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteRecord(sender, e);
        }

        private void updateRecordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "UPDATE";
            setSelectedValues(sender);
        }

        private void insertIntoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "INSERT INTO";
            toolStripButton4_Click(sender, e);
        }

        private void viewScriptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // View/parse script shortcut from right-click
            comboBox1.SelectedIndex = -1;
            toolStripButton3_Click(sender, e);
        }

        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // right-click shortcut to Run from Table View
            comboBox1.SelectedItem = "RUN";
            radioButton1.Checked = true;
            radioButton1_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox2.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox2.SelectedItem;
                button3_Click(sender, e);
            }
        }

        private void toolStripComboBox3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            // right-click shortcut to Run from selected script order and thereafter
            comboBox1.SelectedItem = "RUN";
            radioButton3.Checked = true;
            radioButton3_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox3.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox3.SelectedItem;
                button3_Click(sender, e);
            }
        }

        private void toolStripComboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            // right-click shortcut to Run selected script
            comboBox1.SelectedItem = "RUN";
            radioButton4.Checked = true;
            radioButton4_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox4.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox4.SelectedItem;
                button3_Click(sender, e);
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            // launch SQL Server Management Studio
            System.Diagnostics.Process.Start("SSMS.EXE");
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void lightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flipLightSwitch(true);
        }

        private void darkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            flipLightSwitch(false);
        }

        private void uploadAFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripButton4_Click(sender, e);
        }

        private void restartApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button4_Click(sender, e);
        }

        private void editScriptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "UPDATE";
            setSelectedValues(sender);
        }

        private void deleteScriptToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteRecord(sender, e);
        }


        private void deleteRecord(object sender, EventArgs e)
        {
            int selected = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);

            if (selected < 1)
            {
                MessageBox.Show("You must select a record", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

           // contextMenuStrip Delete Row option
            comboBox1.SelectedIndex = -1;
            comboBox5.SelectedItem = "BRMQATONTDBS07"; // prepare to establish the connection to QAT 7 where table exists

            DialogResult dr = MessageBox.Show("Deleting this record will also delete it from the database.\nAre you sure you want to delete this record?", "WARNING!", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {
                string id = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();

                using (conn = initQatConnection())
                {
                    string query = "DELETE FROM dbo.refresh_testing WHERE id = " + id;

                    try
                    {
                        using (cmd = new SqlCommand(query, conn))
                        {
                            conn.Open();
                            int result = cmd.ExecuteNonQuery();

                            // Check error
                            if (result < 0)
                                toExecLog("Deletion failed.");
                            else
                            {
                                toolStripLabel4.Visible = false;
                                toolStripLabel3.Text = "Record deleted";
                                toolStripLabel3.Visible = true;
                                toExecLog("Record deleted");
                                toolStrip3.Update();
                            }
                            loadData();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + "\n\nPossible Reason: incorrect database connection. Make sure a connection is established in the environment where the table is located.", "Delete Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void insertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            comboBox1.SelectedItem = "INSERT INTO";
            toolStripButton4_Click(sender, e);
        }


        private void toolStripComboBox5_Click(object sender, EventArgs e)
        {
            // right-click shortcut to Run from Table View
            comboBox1.SelectedItem = "RUN";
            radioButton1.Checked = true;
            radioButton1_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox2.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox2.SelectedItem;
                button3_Click(sender, e);
            }
        }


        private void toolStripComboBox6_Click(object sender, EventArgs e)
        {
            // right-click shortcut to Run from selected script order and thereafter
            comboBox1.SelectedItem = "RUN";
            radioButton3.Checked = true;
            radioButton3_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox3.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox3.SelectedItem;
                button3_Click(sender, e);
            }


        }

        private void toolStripComboBox7_Click(object sender, EventArgs e)
        {
            // right-click shortcut to Run selected script
            comboBox1.SelectedItem = "RUN";
            radioButton4.Checked = true;
            radioButton4_CheckedChanged(sender, e);
            contextMenuStrip1.Close();

            if (toolStripComboBox4.SelectedIndex != -1)
            {
                comboBox5.SelectedItem = toolStripComboBox4.SelectedItem;
                button3_Click(sender, e);
            }
        }


        // END: ---------------------------------------------- Menu and contextMenuStrip Items' events ----------------------------------------------
    }
}

