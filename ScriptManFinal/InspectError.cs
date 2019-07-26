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
        public bool colorScheme;

        public InspectError()
        {
            InitializeComponent();
        }
        

        private void Form2_Load(object sender, EventArgs e)
        {
            // color scheme settings
            // light
            if (colorScheme)
            {
                label1.ForeColor = Color.Black;
                label2.ForeColor = Color.Black;
                label3.ForeColor = Color.Black;
                label4.ForeColor = Color.DarkOrange;
                label5.ForeColor = Color.DarkOrange;
                label6.ForeColor = Color.DarkOrange;
                richTextBox1.BackColor = Color.White;
                this.BackColor = Color.LightGray;
                button1.BackColor = Color.DarkOrange;
                button1.ForeColor = Color.White;
                button3.BackColor = Color.DarkOrange;
                button3.ForeColor = Color.White;
                gradientPanel3.ColorLeft = Color.White;
                gradientPanel3.ColorRight = Color.DarkOrange;
                gradientPanel2.ColorLeft = Color.DarkOrange;
                gradientPanel2.ColorRight = Color.White;
            }
            // dark
            else
            {
                label1.ForeColor = Color.Gainsboro;
                label2.ForeColor = Color.Gainsboro;
                label3.ForeColor = Color.Gainsboro;
                label4.ForeColor = Color.Aqua;
                label5.ForeColor = Color.Aqua;
                label6.ForeColor = Color.Aqua;
                richTextBox1.BackColor = Color.LightGray;
                this.BackColor = Color.FromArgb(64, 64, 64);
                button1.BackColor = Color.MediumPurple;
                button1.ForeColor = Color.Indigo;
                button3.BackColor = Color.MediumPurple;
                button3.ForeColor = Color.Indigo;
                gradientPanel3.ColorLeft = Color.Cyan;
                gradientPanel3.ColorRight = Color.Indigo;
                gradientPanel2.ColorLeft = Color.Indigo;
                gradientPanel2.ColorRight = Color.Cyan;
            }

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
