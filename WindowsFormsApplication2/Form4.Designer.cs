using System.Data.OleDb;
namespace WindowsFormsApplication2
{
    partial class Form4
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form4));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.CargarClientes = new System.Windows.Forms.ToolStripMenuItem();
            this.importarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarDesdeExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarDesdeTxtToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.seleccioneSeparadorToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.importarDesdeXMLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importarDesdeJSONToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dataSet1 = new System.Data.DataSet();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 51);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(618, 318);
            this.dataGridView1.TabIndex = 0;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.CargarClientes,
            this.importarToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(642, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // CargarClientes
            // 
            this.CargarClientes.Name = "CargarClientes";
            this.CargarClientes.Size = new System.Drawing.Size(54, 20);
            this.CargarClientes.Text = "Cargar";
            this.CargarClientes.Click += new System.EventHandler(this.CargarClientes_Click);
            // 
            // importarToolStripMenuItem
            // 
            this.importarToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importarDesdeExcelToolStripMenuItem,
            this.importarDesdeTxtToolStripMenuItem,
            this.importarDesdeXMLToolStripMenuItem,
            this.importarDesdeJSONToolStripMenuItem});
            this.importarToolStripMenuItem.Name = "importarToolStripMenuItem";
            this.importarToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.importarToolStripMenuItem.Text = "Importar";
            // 
            // importarDesdeExcelToolStripMenuItem
            // 
            this.importarDesdeExcelToolStripMenuItem.Name = "importarDesdeExcelToolStripMenuItem";
            this.importarDesdeExcelToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.importarDesdeExcelToolStripMenuItem.Text = "Importar desde Excel";
            this.importarDesdeExcelToolStripMenuItem.Click += new System.EventHandler(this.importarDesdeExcelToolStripMenuItem_Click);
            // 
            // importarDesdeTxtToolStripMenuItem
            // 
            this.importarDesdeTxtToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.seleccioneSeparadorToolStripMenuItem,
            this.toolStripSeparator1,
            this.toolStripMenuItem1,
            this.toolStripMenuItem3,
            this.toolStripMenuItem4});
            this.importarDesdeTxtToolStripMenuItem.Name = "importarDesdeTxtToolStripMenuItem";
            this.importarDesdeTxtToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.importarDesdeTxtToolStripMenuItem.Text = "Importar desde Txt / Csv";
            // 
            // seleccioneSeparadorToolStripMenuItem
            // 
            this.seleccioneSeparadorToolStripMenuItem.Enabled = false;
            this.seleccioneSeparadorToolStripMenuItem.Name = "seleccioneSeparadorToolStripMenuItem";
            this.seleccioneSeparadorToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.seleccioneSeparadorToolStripMenuItem.Text = "Seleccione Separador";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(183, 6);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(186, 22);
            this.toolStripMenuItem1.Text = "\" | \"";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(186, 22);
            this.toolStripMenuItem3.Text = "\" , \"";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(186, 22);
            this.toolStripMenuItem4.Text = "\" ; \"";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // importarDesdeXMLToolStripMenuItem
            // 
            this.importarDesdeXMLToolStripMenuItem.Name = "importarDesdeXMLToolStripMenuItem";
            this.importarDesdeXMLToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.importarDesdeXMLToolStripMenuItem.Text = "Importar desde XML";
            this.importarDesdeXMLToolStripMenuItem.Click += new System.EventHandler(this.importarDesdeXMLToolStripMenuItem_Click);
            // 
            // importarDesdeJSONToolStripMenuItem
            // 
            this.importarDesdeJSONToolStripMenuItem.Name = "importarDesdeJSONToolStripMenuItem";
            this.importarDesdeJSONToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.importarDesdeJSONToolStripMenuItem.Text = "Importar desde JSON";
            this.importarDesdeJSONToolStripMenuItem.Click += new System.EventHandler(this.importarDesdeJSONToolStripMenuItem_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(555, 384);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Volver";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 419);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.dataGridView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form4";
            this.Text = "Importar";
            this.Load += new System.EventHandler(this.Form4_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem CargarClientes;
        private System.Windows.Forms.ToolStripMenuItem importarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarDesdeExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importarDesdeTxtToolStripMenuItem;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem seleccioneSeparadorToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem importarDesdeXMLToolStripMenuItem;
        private System.Data.DataSet dataSet1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem importarDesdeJSONToolStripMenuItem;
    }
}