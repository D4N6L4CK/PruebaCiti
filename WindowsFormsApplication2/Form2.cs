using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;

namespace WindowsFormsApplication2
{
    public partial class Form2 : Form
    {
        public string name;
        public Form2(string Usuario)
        {
            name = Usuario;
            InitializeComponent();
        }

        
        private void BtnSalir_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 fr = new Form1();
            fr.Show();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            
        }
        
        private void ultraButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form3 fr = new Form3(name);
            fr.Show();
        }

       

       private void Form2_Load(object sender, EventArgs e)
        {
            String Session = "Usuario: " + name;
            label1.Text = Session;

            
        }

               

        private void CargarClientes_Click(object sender, EventArgs e)
        {
            this.Close();
            Form4 F4 = new Form4(name);
            F4.Show();
        }

        private void exportarCLientesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();
            ExportForm ExpForm = new ExportForm();
            ExpForm.Show();
        }
    }
}
