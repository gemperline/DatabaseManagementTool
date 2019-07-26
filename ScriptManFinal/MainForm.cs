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
using ScriptManFinal.RMDataSet8TableAdapters;
using Microsoft.VisualBasic;
using System.Threading;

// TO DO: improve runtime and efficiency by closing connections when establishing a new one, after execution, 
//            restarts should kill connections and release memory

namespace ScriptManFinal
{
    public partial class MainForm : Form
    {
        SqlConnection conn;
        SqlCommand cmd;
        public static Stream myStream;
        private static string actionSelected = "";
        string selectedID = "";
        DataTable dt;
        BackgroundWorker worker;
        public static Color Yellow;
        private int currentIndex;      // denotes where a RUN sequence stops and proceeds
        private bool proceed;
        private bool ignoreAll;     // ignore all errors for RUN sequence
        private int placeholder;
        private bool lightSwitch = false;
      
        public MainForm()
        {
            InitializeComponent();
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;

            worker.ProgressChanged += Worker_ProgressChanged;
            worker.DoWork += Worker_DoWork;
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // TODO: This line of code loads data into the 'rMDataSet.refresh_testing' table. You can move, or remove it, as needed.
                this.refresh_testingTableAdapter1.Fill(this.rMDataSet8.refresh_testing);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\nTry restarting the application.", "Connection Error");
            }
            setInitState();
        }

        
        private void setInitState()
        {
            toExecSummary("\n------------------------------------------------------------------------------" + 
                              "\nInitializing states...");

            // initial/refresh states and values
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
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 1;
            progressBar1.Step = 1;
            progressBar1.Value = 0;
            radioButton1.Visible = false;
            radioButton3.Visible = false;
            radioButton4.Visible = false;
            textBox1.KeyDown += new KeyEventHandler(execute_KeyDown);
            textBox3.KeyDown += new KeyEventHandler(execute_KeyDown);
            toolStripTextBox1.KeyUp += new KeyEventHandler(toolStripTextBox1_KeyUp);
            progressBar1.Value = 0;
            label3.Visible = false;
            tabControl2.SelectedIndex = 0;

            if (comboBox1.Items.Contains("UPDATE & RESUME"))
                comboBox1.Items.Remove("UPDATE & RESUME");
            if (!comboBox1.Items.Contains("INSERT INTO"))
                comboBox1.Items.Add("INSERT INTO");
            if (!comboBox1.Items.Contains("RUN"))
                comboBox1.Items.Add("RUN");

            toExecSummary("States initialized");
            loadData();
            dataGridView1.SelectedRows[0].Selected = false;
        }

        private void button1_Click(object sender, EventArgs e)
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



        private void button5_Click(object sender, EventArgs e)
        {
            // close button
            Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
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

                if (!comboBox5.Items.Contains("QUAT_ALL"))
                    comboBox5.Items.Add("QUAT_ALL");
            }
            else if (actionSelected == "UPDATE")
            {
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

                if (!comboBox5.Items.Contains("QUAT_ALL"))
                    comboBox5.Items.Add("QUAT_ALL");
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

                if (comboBox5.Items.Contains("QUAT_ALL"))
                    comboBox5.Items.Remove("QUAT_ALL");
            }

            // when UPDATE & RESUME is selected
            else if (actionSelected == "UPDATE & RESUME")
            {
                textBox5.Text = dataGridView1.SelectedRows[0].Cells[0].ToString();
                label10.Enabled = true;
                label14.Visible = true;
                label14.Enabled = false;
                textBox1.Enabled = false;
                label17.Enabled = true;
                label17.Visible = true;
                label12.Text = "Viewing:";
                comboBox2.Enabled = false;
                comboBox2.Visible = true;
                label11.Enabled = true;
                label11.Visible = true;
                textBox3.Enabled = false;
                textBox3.Visible = true;
                label13.Enabled = true;
                comboBox5.Visible = true;
                textBox2.Text = "AUTO";
                label17.Text = "Run Environment:";
                comboBox5.Enabled = true;
                textBox5.Visible = true;
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
            }

            // no selection
            else if (comboBox1.SelectedIndex == -1)
            {
                label12.Text = "Viewing:";
            }

            if (label5.Enabled)
                label4.Enabled = true;
            else
                label4.Enabled = false;
        }



        private void button4_Click(object sender, EventArgs e)
        {
            toExecSummary("Restarting...");
            {
                if (conn != null && conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                setInitState();
            }
        }



        // CAUTION! ------------------------------------------------------- EXECUTION STATEMENTS --------------------------------------------------------------------------------------
        private void button3_Click(object sender, EventArgs e)
        {   // EXECUTE button 

            double progressPercent = 0.0;
            proceed = true;     // control variable that stops RUN sequence upon user's error response

            // determine number of records in table
            int numRows = dataGridView1.Rows.Count;
            toExecSummary("Number of Records in View: " + numRows);
            int lastIndex = numRows - 1;

            // show Execution Summary tab
            tabControl1.SelectedIndex = 1;

            if (conn.State == ConnectionState.Open)
                conn.Close();

            if (fieldCheck(sender, e) == false)
                return;
                
            if (actionSelected == "RUN")
            {
                // if "From Table View" is selected
                if ((radioButton1.Checked == true) && (comboBox5.SelectedIndex != -1))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        // select current row
                        dataGridView1.Rows[row.Index].Selected = true;
                        currentIndex = row.Index;

                        if (row.Index.Equals(numRows))
                        {
                            // loop exceeded range
                            break;
                        }
                        // run all scripts matching selected environment or where environment is QUAT_ALL
                        else if (proceed == true && (row.Cells[3].Value.ToString() == comboBox5.SelectedItem.ToString() || row.Cells[3].Value.ToString() == "QUAT_ALL"))
                        {
                            runScript(sender, row, e);

                            progressBar1.PerformStep();
                            progressPercent = progressBar1.Value;
                            label20.Text = progressBar1.Value.ToString() + "%";
                        }

                        if (proceed == false)
                        {
                            break;
                        }

                        // if this row is the last row of the displayed records, restart the application
                        if (row.Index == dataGridView1.Rows.GetLastRow(DataGridViewElementStates.Displayed))
                        {
                            toExecSummary("Run sequence complete.");
                            button4_Click(sender, e);
                        }

                    }                 
                }
                // Start At Script Order is selected 
                else if (radioButton3.Checked == true && (comboBox5.SelectedIndex != -1))
                {

                    int startingIndex = dataGridView1.SelectedRows[0].Index;

                    // loop to run scripts within given script order range
                    for (int i = startingIndex; i < numRows; i++)
                    {
                        // select current row
                        dataGridView1.Rows[i].Selected = true;
                        toExecSummary("CURRENT INDEX: " + i + ", LAST INDEX: " + lastIndex);

                        // set the selected row to the current index
                        dataGridView1.SelectedRows[0].Index.Equals(i);

                        // run all scripts matching selected environment or where environment is QUAT_ALL
                        if (proceed == true && (dataGridView1.SelectedRows[0].Cells[3].Value.ToString() == comboBox5.SelectedItem.ToString() || dataGridView1.SelectedRows[0].Cells[3].Value.ToString() == "QUAT_ALL"))
                        {
                            runScript(sender, dataGridView1.SelectedRows[0], e);

                            progressBar1.PerformStep();
                            progressPercent = progressBar1.Value;
                            label20.Text = progressBar1.Value.ToString() + "%";
                        }

                        if (proceed == false)
                        {
                            break;
                        }

                        // if this row is the last row of the displayed records, restart the application
                        if (i == dataGridView1.Rows.GetLastRow(DataGridViewElementStates.Displayed))
                        {
                            toExecSummary("Last index reached.");
                            button4_Click(sender, e);
                        }
                    }
                }
                // Run Selected Script is selected 
                else if ((comboBox5.SelectedIndex != -1) && (radioButton4.Checked == true))
                {
                    DataGridViewRow row = dataGridView1.SelectedRows[0];
                    runScript(sender, row, e);
                    progressBar1.PerformStep();
                    progressPercent = progressBar1.Value;
                    toExecSummary("Progress: " + progressPercent);
                    label20.Text = progressBar1.Value.ToString() + "%";
                    setInitState();
                }
                else
                    fieldCheck(sender, e);
            }
            else if (comboBox1.SelectedIndex == -1 || string.IsNullOrWhiteSpace(textBox1.Text) || comboBox2.SelectedIndex == -1 || comboBox3.SelectedIndex == -1 || string.IsNullOrWhiteSpace(textBox3.Text) || string.IsNullOrWhiteSpace(richTextBox1.Text))
            {
                toExecSummary("Incomplete field(s). Execution stopped.");
                fieldCheck(sender, e);
            }
            else if (actionSelected == "UPDATE" && proceed == true)
            {
                proceed = false;
                updateRecord(sender, e);
            }
            else if (actionSelected == "INSERT INTO")
            {
                proceed = false;
                insertIntoTable(sender , e);
            }
            else if (actionSelected == "UPDATE & RESUME")
            {
                proceed = true;
                updateRecord(sender, e);
                actionSelected = "RUN";
                radioButton3.Checked = true;
                button3_Click(sender, e);
            }


            toExecSummary("ProgressBar Value: " + progressBar1.Value);

            // reset progress bar value
            if (progressBar1.Value == progressBar1.Maximum)
            {
                label19.Text = "Process completed.";

            }
        }



        // ------------------------------------------------Update existing script--------------------------------------------------
        private void updateRecord(object sender, EventArgs e)
        {
            using (conn = initQatConnection())
            {
                string query = "UPDATE dbo.refresh_testing SET Script=@script, script_function=@script_function, script_environment=@script_environment, script_description=@script_description, script_order=@script_order WHERE ID = @id";
                using (cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("id", selectedID);
                    cmd.Parameters.AddWithValue("script", richTextBox1.Text);
                    cmd.Parameters.AddWithValue("script_function", comboBox2.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("script_environment", comboBox3.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("script_description", textBox3.Text);
                    cmd.Parameters.AddWithValue("script_order", Convert.ToDecimal(textBox1.Text));

                    toExecSummary("Updating record...");
                    conn.Open();
                    int result = cmd.ExecuteNonQuery();

                    // Check error
                    if (result < 0)
                    {
                        toExecSummary("Could not update script. Check script for errors. Result = " + result);
                    }
                    else
                        toExecSummary("Script updated successfully.");
                    toExecSummary("\n------------------------------------------------------------------------------");
                }
            }
            if (proceed == false)
                button4_Click(sender, e); // restart
        }



        // ------------------------------------------------Insert a new script--------------------------------------------------
        private void insertIntoTable(object sender, EventArgs e)
        {
            using (conn = initQatConnection())
            {
                string query = "INSERT INTO dbo.refresh_testing (Script, script_function, script_environment, script_description, script_order) VALUES (@Script, @script_function, @script_environment, @script_description, @script_order)";
                using (cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Script", richTextBox1.Text);
                    cmd.Parameters.AddWithValue("@script_function", comboBox2.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@script_environment", comboBox3.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@script_description", textBox3.Text);
                    cmd.Parameters.AddWithValue("@script_order", Convert.ToDecimal(textBox1.Text));

                    toExecSummary("Inserting record...");
                    conn.Open();
                    int result = cmd.ExecuteNonQuery();

                    // Check error
                    if (result < 0)
                    {
                        toExecSummary("Could not insert script. Check script for errors. Result = " + result);
                    }
                    else
                        toExecSummary("Script inserted successfully.");
                    toExecSummary("\n------------------------------------------------------------------------------");
                }
            }
            if (proceed == false)
                button4_Click(sender, e); // restart
        }



        // ---------------------------------------------------------Run script(s)------------------------------------------------------------
        private void runScript(object sender, DataGridViewRow currentRow, EventArgs e)
        {
            // prompt user if selected script field is blank upon execution, offer to parse from selection
            if (string.IsNullOrWhiteSpace(richTextBox1.Text) && radioButton4.Checked == true)
            {
                DialogResult diRes = MessageBox.Show("A script must be loaded. Would you like to use the selected script?", "ATTENTION", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                if (diRes == DialogResult.Yes)
                {
                    DataGridView dgv = sender as DataGridView;
                    if (dgv != null && dgv.SelectedRows.Count > 0)
                    {
                        // set selected Script query to view tab
                        richTextBox1.Text = currentRow.Cells[5].Value.ToString();
                    }
                }
                else
                    return;
            }
            else
            {
                string query = currentRow.Cells[5].Value.ToString();
                toExecSummary("\n____________________________________________" +
                              "\nRunning Script Order: " + currentRow.Cells[0].Value.ToString() + ", ID: " + currentRow.Cells[1].Value.ToString() +
                              "\n------------------------------------------------------------------------------");
                toExecSummary("Current Index: " + currentRow.Index);

                try
                {
                    using (conn = initQatConnection())
                    {
                        try
                        {
                            using (cmd = new SqlCommand(query, conn))
                            {
                                conn.Open();
                                toExecSummary("Connected to " + comboBox5.SelectedItem.ToString());                            
                                try 
                                {
                                    int result = cmd.ExecuteNonQuery();
                                    toExecSummary("Run result = " + result);
                                }            
                                catch ( SqlException ex)
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
                                MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (SqlException ex)
                {
                    logError(sender, ex, currentRow, e);
                }
            }
        }
        //  --------------------------------------------------------------- END EXECUTION STATEMENTS --------------------------------------------------------------------------------------
  


        private void logError(object sender, SqlException exception, DataGridViewRow currentScriptRow, EventArgs e)
        {
            string currQAT = null;
            // if selected Run Environment is not BRMQATONTDBS07, force conncection for error logging (dbo.refresh_error_log is in BRMQATONTDBS07), then return to previous selection
            if (comboBox5.SelectedItem.ToString() != "BRMQATONTDBS07")
            {
                currQAT = comboBox5.SelectedItem.ToString();
                toExecSummary("Current QAT connection: " + currQAT);
                comboBox5.SelectedItem = "BRMQATONTDBS07";
                toExecSummary("Connecting to " + comboBox5.SelectedItem.ToString() + " for error logging...");
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
                        toExecSummary("Logging error...");
                        int result = cmd.ExecuteNonQuery();

                        // Check error
                        if (result < 0)
                            toExecSummary("Error logging failed. Result = " + result);
                        else
                            toExecSummary("Error logged. Result = " + result);
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show(ex.Message);
                        toExecSummary(ex.Message);
                    }     
                }
            }
            refreshErrorLog();
            if (currQAT != "BRMQATONTDBS07" && currQAT != null)
            {
                comboBox5.SelectedItem = currQAT;
                toExecSummary("Connection restored to: " + comboBox5.SelectedItem.ToString());
            }
            

            // an Ignore All response from error dialog will prevent future error dialogs for this run session
            toExecSummary("Ignore All status: " + ignoreAll);
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
                        toExecSummary("Ignoring error");
                    else if (response == "ignore all")
                    {
                        ignoreAll = true;
                        toExecSummary("Ignoring all errors for this run session\n\t -errors will still be logged");
                    }
                    else if (response == "edit script")
                    {
                        toExecSummary("Entering edit mode...");
                        prepareForUpdateAndResume(sender, e, exception, currentScriptRow);

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
                        toExecSummary(response + "\nAborting...");
                        proceed = false;
                    }
                }
                else
                    MessageBox.Show("The error log is empty");
            }
            Console.WriteLine("Ignore all is true");           
        }



        private void prepareForUpdateAndResume(object sender, EventArgs e, SqlException exception, DataGridViewRow currentScriptRow)
        {
            toExecSummary("Current Index: " + currentScriptRow.Index);

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
            toExecSummary("Execution stopped at row: " + currentScriptRow.Index);
            comboBox1.SelectedIndex = 1;
            textBox5.Text = currentScriptRow.Index.ToString();
            radioButton3.Checked = true;
            setSelectedValues(sender);
        }


        private void selectLastErrorLogged()
        {
            // select most recent (last) error record         
            int lastRow = dataGridView2.Rows.Count - 1;
            dataGridView2.Rows[lastRow].Selected = true;
            dataGridView2.FirstDisplayedScrollingRowIndex = lastRow;
        }



        private static void highlightLine(RichTextBox rtb1, int index, Color color)
        {
            var lines = rtb1.Lines;
            if (index < 0 || index >= lines.Length)
                return;
            var start = rtb1.GetFirstCharIndexFromLine(index);
            var length = lines[index].Length;
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
                toExecSummary("Opening error dialog...");
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
                toExecSummary("Your response: " + errorOption);
                f2.Close();
                return errorOption;
            }
            string failure = "Failed to open error dialog: selected error record is null or nothing was selected";
            return failure;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            // LOCK / UNLOCK BUTTON 
            if (richTextBox1.Enabled == false)
            {
                richTextBox1.Enabled = true;
                button2.Text = "LOCK";
            }
            else
            {
                richTextBox1.Enabled = false;
                button2.Text = "UNLOCK";
            }
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
                        toExecSummary("Loading table...");
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
            Manoli.Utils.CSharpFormat.TsqlFormat clrz = new Manoli.Utils.CSharpFormat.TsqlFormat();
            richTextBox1.Rtf = clrz.FormatString(richTextBox1.Text);
        }



        private void button6_Click(object sender, EventArgs e)
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



        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // allows for digits, decimal, and backspace / delete entries
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != 46)
            {
                e.Handled = true;
                return;
            }

            // allows for only one decimal to be entered
            if ( e.KeyChar == 46)
            {
                if ((sender as TextBox).Text.IndexOf(e.KeyChar) != - 1)
                e.Handled = true;
            }

            // allows only two digits after decimal
            if (Regex.IsMatch(textBox1.Text, @"\.\d\d"))
            {
                e.Handled = true;
            }
        }



        private void decryptAccess()
        {
            // calls decrypt function in DB "master" to check user's permission to run scripts on user selected QAT
            // decrypt function checks to see if system user name is found in table for corresponding QAT connection. If matching record exists, 1 is returned. 
            object result = 0;
            
            try
            {
                if (conn != null && conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    toExecSummary("Connection was open, now closed");
                }

                using (conn = initQatConnection())
                {
                    string query = @"DECLARE @hasAccess INT = 0, @sql NVARCHAR(MAX) BEGIN TRY set @sql = 'SELECT @hasAccess = [master].[dbo].[fn_get_user_access] (SUSER_SNAME() , ''OnTrakProd'')' EXEC sp_executeSQL @sql, N'@hasAccess int output', @hasAccess output SELECT @hasAccess END TRY BEGIN CATCH set @sql = 'SELECT @hasAccess = [master].[dbo].[fn_get_user_access] ( SUSER_SNAME() , ''OnTrakProd'', NULL)' EXEC sp_executeSQL @sql, N'@hasAccess int output', @hasAccess output SELECT @hasAccess END CATCH";
                    SqlCommand cmd = new SqlCommand(query, conn);
                    conn.Open();
                    cmd = new SqlCommand(query, conn);
                    result = cmd.ExecuteScalar();

                    // Decrypt Access response
                    string access = "";
                    if ((int)result != 1)
                    {
                        label3.Text = "CAUTION: You do not have Decrypt Access in database " + comboBox5.SelectedItem.ToString();
                        label3.Visible = true;
                        access = "Denied";
                    }
                    else
                    {
                        access = "Granted";
                        label3.Visible = false;
                    }
                    toExecSummary("Decrypt Access: " + access + " @ " + comboBox5.SelectedItem.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error openning connection: " + ex.Message);
            }
        }



        private SqlConnection initQatConnection()
        {
            // this function initializes and then returns a new SqlConnection
            // if the action selected is RUN, this function will establish a connection based on the User's Run Environment selection 
                 // otherwise it will connect to QAT7 for all other purposes

            if (conn != null && conn.State == ConnectionState.Open)
                conn.Close();

            // RM connecion strings 
            
            // if RUN is selected along with a Run Environment from combobox5, connect to the selected Run Environment
            if (actionSelected == "RUN" && comboBox5.SelectedIndex != -1) 
            {
                // no selection
                if (comboBox5.SelectedIndex == -1)
                    return null;
                // brmqatontdbs02 RM
                else if (comboBox5.SelectedIndex == 0)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs02;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs03 RM
                else if (comboBox5.SelectedIndex == 1)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs03;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs04 RM
                else if (comboBox5.SelectedIndex == 2)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs04;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs05 RM
                else if (comboBox5.SelectedIndex == 3)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs05;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs06 RM
                else if (comboBox5.SelectedIndex == 4)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs06;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs07 RM
                else if (comboBox5.SelectedIndex == 5)
                {
                    conn = new SqlConnection(@"Data Source=Brmqatontdbs07;Initial Catalog=RM;Integrated Security=True");
                }
                // brmqatontdbs11 RM
                else if (comboBox5.SelectedIndex == 6)
                {
                    conn = new SqlConnection(@"Data Source=brmqatontdbs11;Initial Catalog=RM;Integrated Security=True");
                }
                else
                {
                    toExecSummary("The selected QAT Environment " + comboBox5.GetItemText(comboBox5.SelectedItem) + " does not have a defined connection.");
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
                         toExecSummary("Loading table...");
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

                if (actionSelected == "UPDATE" || (actionSelected == "RUN" && radioButton4.Checked == true))
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
                        toExecSummary("Deletion failed.");
                    else
                        label19.Text = "Deletion successful.";

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
                        toExecSummary("Loading error log...");
                        da.Fill(dt);
                        dataGridView2.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error establishing connection to table", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            toExecSummary("Load complete" +
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
                    toExecSummary("Inspecting error...");
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
                        int scriptID = Convert.ToInt32(row.Cells[1].Value);
                        tabControl1.SelectedIndex = 0;
                        richTextBox1.Enabled = true;

                        // search the script table for matching script
                        DataGridView grid = dataGridView1 as DataGridView;
                        for (int i = 0; i < grid.Rows.Count - 1; i++)
                        {
                            if (scriptID == Convert.ToInt32(grid.Rows[i].Cells[1].Value))
                            {
                                richTextBox1.Text = grid.Rows[i].Cells[5].Value.ToString();
                                colorizeScript();
                                highlightLine(richTextBox1, 33, Yellow);
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Please select a valid error record");
            }
        }




        private Task processData(List<string> list, IProgress<ProgressReport> progress)
        {
            int index = 1;
            int totalProcess = list.Count;
            var progressReport = new ProgressReport();
            return Task.Run(() =>
            {
                for (int i = 0; i < totalProcess; i++)
                {
                    progressReport.percentComplete = index++ * 100 / totalProcess;
                    progress.Report(progressReport);
                    Thread.Sleep(10);//used to simulate length of operation
                }
            });
        }



        private void toExecSummary(string output)
        {
            // sends input to Execution Summary rich text box
            richTextBox2.AppendText("\r" + output);
            richTextBox2.ScrollToCaret();
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
                button6_Click(sender, e);
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
            comboBox5.SelectedIndex = -1;
            richTextBox1.Text = null;

            if (richTextBox1.Enabled == true)
                button2_Click(sender, e);
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
                comboBox5.Visible = false;
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
                label3.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label16.ForeColor = Color.Black;
                label17.ForeColor = Color.Black;
                label19.ForeColor = Color.Black;
                label20.ForeColor = Color.Black;
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
                button1.BackColor = Color.Orange;
                button1.ForeColor = Color.Black;
                button6.BackColor = Color.Moccasin;
                button6.ForeColor = Color.Black;
                button2.BackColor = Color.Moccasin;
                button2.ForeColor = Color.Black;
                button3.BackColor = Color.ForestGreen;
                button4.BackColor = Color.Gold;
                button5.BackColor = Color.Firebrick;
                this.ForeColor = Color.Black;
                gradientPanel1.ColorLeft = Color.White;
                gradientPanel1.ColorRight= Color.DarkOrange;
                gradientPanel2.ColorLeft = Color.DarkOrange;
                gradientPanel2.ColorRight = Color.White;
                richTextBox2.BackColor = Color.White;
                richTextBox2.ForeColor = Color.Black;
                radioButton1.ForeColor = Color.Black;
                radioButton3.ForeColor = Color.Black;
                radioButton4.ForeColor = Color.Black;             
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
                label3.ForeColor = Color.Gainsboro;
                label9.ForeColor = Color.Gainsboro;
                label16.ForeColor = Color.Gainsboro;
                label17.ForeColor = Color.Gainsboro;
                label19.ForeColor = Color.Gainsboro;
                label20.ForeColor = Color.Gainsboro;
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
                button1.BackColor = Color.RoyalBlue;
                button1.ForeColor = Color.White;
                button6.BackColor = Color.MediumTurquoise;
                button6.ForeColor = Color.Black;
                button2.BackColor = Color.LimeGreen;
                button2.ForeColor = Color.White;
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


    }
}

