namespace schedule
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.создатьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.создатьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.печатьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.правкаToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьПредметToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьПредметToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.предметыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.удалитьУчителяToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.копироватьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.вставитьToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.копироватьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.вставитьToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.предпросмотрРасписанияToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            resources.ApplyResources(this.menuStrip1, "menuStrip1");
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.создатьToolStripMenuItem,
            this.правкаToolStripMenuItem,
            this.оПрограммеToolStripMenuItem});
            this.menuStrip1.Name = "menuStrip1";
            // 
            // создатьToolStripMenuItem
            // 
            resources.ApplyResources(this.создатьToolStripMenuItem, "создатьToolStripMenuItem");
            this.создатьToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.создатьToolStripMenuItem1,
            this.сохранитьToolStripMenuItem,
            this.печатьToolStripMenuItem,
            this.предпросмотрРасписанияToolStripMenuItem});
            this.создатьToolStripMenuItem.Name = "создатьToolStripMenuItem";
            // 
            // создатьToolStripMenuItem1
            // 
            resources.ApplyResources(this.создатьToolStripMenuItem1, "создатьToolStripMenuItem1");
            this.создатьToolStripMenuItem1.Name = "создатьToolStripMenuItem1";
            this.создатьToolStripMenuItem1.Click += new System.EventHandler(this.создатьToolStripMenuItem1_Click);
            // 
            // сохранитьToolStripMenuItem
            // 
            resources.ApplyResources(this.сохранитьToolStripMenuItem, "сохранитьToolStripMenuItem");
            this.сохранитьToolStripMenuItem.Name = "сохранитьToolStripMenuItem";
            this.сохранитьToolStripMenuItem.Click += new System.EventHandler(this.сохранитьToolStripMenuItem_Click_1);
            // 
            // печатьToolStripMenuItem
            // 
            resources.ApplyResources(this.печатьToolStripMenuItem, "печатьToolStripMenuItem");
            this.печатьToolStripMenuItem.Name = "печатьToolStripMenuItem";
            this.печатьToolStripMenuItem.Click += new System.EventHandler(this.печатьToolStripMenuItem_Click);
            // 
            // правкаToolStripMenuItem
            // 
            resources.ApplyResources(this.правкаToolStripMenuItem, "правкаToolStripMenuItem");
            this.правкаToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.удалитьПредметToolStripMenuItem,
            this.добавитьПредметToolStripMenuItem,
            this.предметыToolStripMenuItem,
            this.удалитьУчителяToolStripMenuItem});
            this.правкаToolStripMenuItem.Name = "правкаToolStripMenuItem";
            // 
            // удалитьПредметToolStripMenuItem
            // 
            resources.ApplyResources(this.удалитьПредметToolStripMenuItem, "удалитьПредметToolStripMenuItem");
            this.удалитьПредметToolStripMenuItem.Name = "удалитьПредметToolStripMenuItem";
            this.удалитьПредметToolStripMenuItem.Click += new System.EventHandler(this.удалитьПредметToolStripMenuItem_Click);
            // 
            // добавитьПредметToolStripMenuItem
            // 
            resources.ApplyResources(this.добавитьПредметToolStripMenuItem, "добавитьПредметToolStripMenuItem");
            this.добавитьПредметToolStripMenuItem.Name = "добавитьПредметToolStripMenuItem";
            this.добавитьПредметToolStripMenuItem.Click += new System.EventHandler(this.добавитьПредметToolStripMenuItem_Click);
            // 
            // предметыToolStripMenuItem
            // 
            resources.ApplyResources(this.предметыToolStripMenuItem, "предметыToolStripMenuItem");
            this.предметыToolStripMenuItem.Name = "предметыToolStripMenuItem";
            this.предметыToolStripMenuItem.Click += new System.EventHandler(this.предметыToolStripMenuItem_Click);
            // 
            // удалитьУчителяToolStripMenuItem
            // 
            resources.ApplyResources(this.удалитьУчителяToolStripMenuItem, "удалитьУчителяToolStripMenuItem");
            this.удалитьУчителяToolStripMenuItem.Name = "удалитьУчителяToolStripMenuItem";
            this.удалитьУчителяToolStripMenuItem.Click += new System.EventHandler(this.удалитьУчителяToolStripMenuItem_Click);
            // 
            // оПрограммеToolStripMenuItem
            // 
            resources.ApplyResources(this.оПрограммеToolStripMenuItem, "оПрограммеToolStripMenuItem");
            this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
            // 
            // копироватьToolStripMenuItem
            // 
            resources.ApplyResources(this.копироватьToolStripMenuItem, "копироватьToolStripMenuItem");
            this.копироватьToolStripMenuItem.Name = "копироватьToolStripMenuItem";
            // 
            // вставитьToolStripMenuItem
            // 
            resources.ApplyResources(this.вставитьToolStripMenuItem, "вставитьToolStripMenuItem");
            this.вставитьToolStripMenuItem.Name = "вставитьToolStripMenuItem";
            // 
            // копироватьToolStripMenuItem1
            // 
            resources.ApplyResources(this.копироватьToolStripMenuItem1, "копироватьToolStripMenuItem1");
            this.копироватьToolStripMenuItem1.Name = "копироватьToolStripMenuItem1";
            // 
            // вставитьToolStripMenuItem1
            // 
            resources.ApplyResources(this.вставитьToolStripMenuItem1, "вставитьToolStripMenuItem1");
            this.вставитьToolStripMenuItem1.Name = "вставитьToolStripMenuItem1";
            // 
            // предпросмотрРасписанияToolStripMenuItem
            // 
            resources.ApplyResources(this.предпросмотрРасписанияToolStripMenuItem, "предпросмотрРасписанияToolStripMenuItem");
            this.предпросмотрРасписанияToolStripMenuItem.Name = "предпросмотрРасписанияToolStripMenuItem";
            this.предпросмотрРасписанияToolStripMenuItem.Click += new System.EventHandler(this.предпросмотрРасписанияToolStripMenuItem_Click);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem создатьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem правкаToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem создатьToolStripMenuItem1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ToolStripMenuItem копироватьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem вставитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem копироватьToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem вставитьToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem сохранитьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem печатьToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem добавитьПредметToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem удалитьПредметToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem предметыToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem удалитьУчителяToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem предпросмотрРасписанияToolStripMenuItem;
    }
}

