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


namespace ScriptManFinal
{
    public partial class InspectError : Form
    {
        public string errorMessage;
        public int errorLine;
        public string procedure;
        public string option = null;
        public string scriptOrder;
        public string id;
        public string env;

        public InspectError()
        {
            InitializeComponent();
        }
        

        private void Form2_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(procedure))
                procedure = "N/A";

            label4.Text = scriptOrder;
            label5.Text = id;
            label6.Text = env;
            richTextBox1.Text = "ERROR MESSAGE: " + errorMessage + "\n\nPROCEDURE: " + procedure + "\n\nAT LINE: " + errorLine;
        }



        private void button3_Click(object sender, EventArgs e)
        {
            option = "close";
            Close();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            option = "edit script";
            Close();
        }


    }

}
