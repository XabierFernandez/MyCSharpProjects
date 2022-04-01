using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GUI_PDF2EXCEL
{
    public partial class Form2 : Form 
    {
        public Form2()
        {
            InitializeComponent();
            textBox3.Text = Form1.getCodOper();
            textBox1.Text = Form1.getEmail();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if ((textBox1.Text != "") && (textBox3.Text != ""))
                {
                    var test = new MailAddress(textBox1.Text);
                    Form1.setCodOper(textBox3.Text);
                    Form1.setEmail(textBox1.Text);
                    Form1.saveXml();
                    MessageBox.Show(" Se han guardado los datos");
                }
                else{
                    MessageBox.Show("Los campos estan vacios o incompletos", "Atencion",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch(FormatException ex)
            {
                MessageBox.Show(" Error: Campo de direccion email incorrecto \n\n","Atencion",MessageBoxButtons.OK,MessageBoxIcon.Warning );
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Se va a proceder a limpiar campos\n ¿Quiere continuar?", "Atencion",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                textBox3.Text = "";
                textBox1.Text = "";
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if ((textBox1.Text == "") || (textBox3.Text == ""))
            {
                if (MessageBox.Show("Los campos estan vacios o incompletos\n¿Quieres cerrar ventana?", "Atencion",
                    MessageBoxButtons.YesNo,MessageBoxIcon.Warning) == DialogResult.No)
                {
                    // Cancel the Closing event from closing the form.
                    e.Cancel = true;
                    // Call method to save file...
                }
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
            
        }
    }
}
