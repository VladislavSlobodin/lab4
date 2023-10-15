namespace lab4
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            WordDataGridView = new DataGridView();
            tabControl1 = new TabControl();
            WordTab = new TabPage();
            menuStrip1 = new MenuStrip();
            открытьШаблонToolStripMenuItem = new ToolStripMenuItem();
            сохранитьРезультатToolStripMenuItem = new ToolStripMenuItem();
            ExcelTab = new TabPage();
            ExcelDataGridView = new DataGridView();
            menuStrip2 = new MenuStrip();
            рассчитатьToolStripMenuItem = new ToolStripMenuItem();
            выбратьПутьДляСохраненияToolStripMenuItem = new ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)WordDataGridView).BeginInit();
            tabControl1.SuspendLayout();
            WordTab.SuspendLayout();
            menuStrip1.SuspendLayout();
            ExcelTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)ExcelDataGridView).BeginInit();
            menuStrip2.SuspendLayout();
            SuspendLayout();
            // 
            // WordDataGridView
            // 
            WordDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            WordDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            WordDataGridView.Dock = DockStyle.Fill;
            WordDataGridView.Location = new Point(3, 3);
            WordDataGridView.Name = "WordDataGridView";
            WordDataGridView.RowTemplate.Height = 25;
            WordDataGridView.Size = new Size(786, 416);
            WordDataGridView.TabIndex = 1;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(WordTab);
            tabControl1.Controls.Add(ExcelTab);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(800, 450);
            tabControl1.TabIndex = 3;
            // 
            // WordTab
            // 
            WordTab.Controls.Add(menuStrip1);
            WordTab.Controls.Add(WordDataGridView);
            WordTab.Location = new Point(4, 24);
            WordTab.Name = "WordTab";
            WordTab.Padding = new Padding(3);
            WordTab.Size = new Size(792, 422);
            WordTab.TabIndex = 0;
            WordTab.Text = "Word";
            WordTab.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { открытьШаблонToolStripMenuItem, сохранитьРезультатToolStripMenuItem });
            menuStrip1.Location = new Point(3, 3);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(786, 24);
            menuStrip1.TabIndex = 3;
            menuStrip1.Text = "menuStrip1";
            // 
            // открытьШаблонToolStripMenuItem
            // 
            открытьШаблонToolStripMenuItem.Name = "открытьШаблонToolStripMenuItem";
            открытьШаблонToolStripMenuItem.Size = new Size(114, 20);
            открытьШаблонToolStripMenuItem.Text = "Открыть шаблон";
            // 
            // сохранитьРезультатToolStripMenuItem
            // 
            сохранитьРезультатToolStripMenuItem.Name = "сохранитьРезультатToolStripMenuItem";
            сохранитьРезультатToolStripMenuItem.Size = new Size(134, 20);
            сохранитьРезультатToolStripMenuItem.Text = "Сохранить результат";
            // 
            // ExcelTab
            // 
            ExcelTab.Controls.Add(ExcelDataGridView);
            ExcelTab.Controls.Add(menuStrip2);
            ExcelTab.Location = new Point(4, 24);
            ExcelTab.Name = "ExcelTab";
            ExcelTab.Padding = new Padding(3);
            ExcelTab.Size = new Size(792, 422);
            ExcelTab.TabIndex = 1;
            ExcelTab.Text = "Excel";
            ExcelTab.UseVisualStyleBackColor = true;
            // 
            // ExcelDataGridView
            // 
            ExcelDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            ExcelDataGridView.Dock = DockStyle.Fill;
            ExcelDataGridView.Location = new Point(3, 27);
            ExcelDataGridView.Name = "ExcelDataGridView";
            ExcelDataGridView.RowTemplate.Height = 25;
            ExcelDataGridView.Size = new Size(786, 392);
            ExcelDataGridView.TabIndex = 0;
            // 
            // menuStrip2
            // 
            menuStrip2.Items.AddRange(new ToolStripItem[] { выбратьПутьДляСохраненияToolStripMenuItem, рассчитатьToolStripMenuItem });
            menuStrip2.Location = new Point(3, 3);
            menuStrip2.Name = "menuStrip2";
            menuStrip2.Size = new Size(786, 24);
            menuStrip2.TabIndex = 1;
            menuStrip2.Text = "menuStrip2";
            // 
            // рассчитатьToolStripMenuItem
            // 
            рассчитатьToolStripMenuItem.Name = "рассчитатьToolStripMenuItem";
            рассчитатьToolStripMenuItem.Size = new Size(80, 20);
            рассчитатьToolStripMenuItem.Text = "Рассчитать";
            рассчитатьToolStripMenuItem.Click += рассчитатьToolStripMenuItem_Click;
            // 
            // выбратьПутьДляСохраненияToolStripMenuItem
            // 
            выбратьПутьДляСохраненияToolStripMenuItem.Name = "выбратьПутьДляСохраненияToolStripMenuItem";
            выбратьПутьДляСохраненияToolStripMenuItem.Size = new Size(183, 20);
            выбратьПутьДляСохраненияToolStripMenuItem.Text = "Выбрать путь для сохранения";
            выбратьПутьДляСохраненияToolStripMenuItem.Click += выбратьПутьДляСохраненияToolStripMenuItem_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(tabControl1);
            Name = "Form1";
            ((System.ComponentModel.ISupportInitialize)WordDataGridView).EndInit();
            tabControl1.ResumeLayout(false);
            WordTab.ResumeLayout(false);
            WordTab.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ExcelTab.ResumeLayout(false);
            ExcelTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)ExcelDataGridView).EndInit();
            menuStrip2.ResumeLayout(false);
            menuStrip2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private DataGridView WordDataGridView;
        private TabControl tabControl1;
        private TabPage WordTab;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem открытьШаблонToolStripMenuItem;
        private ToolStripMenuItem сохранитьРезультатToolStripMenuItem;
        private TabPage ExcelTab;
        private DataGridView ExcelDataGridView;
        private MenuStrip menuStrip2;
        private ToolStripMenuItem рассчитатьToolStripMenuItem;
        private ToolStripMenuItem выбратьПутьДляСохраненияToolStripMenuItem;
    }
}