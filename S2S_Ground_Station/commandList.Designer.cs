namespace S2S_Ground_Station
{
    partial class commandList
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
            components = new System.ComponentModel.Container();
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle4 = new DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(commandList));
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            dataGridView1 = new DataGridView();
            command_a = new DataGridViewTextBoxColumn();
            comment_a = new DataGridViewTextBoxColumn();
            remark_a = new DataGridViewTextBoxColumn();
            tabPage2 = new TabPage();
            dataGridView2 = new DataGridView();
            command_B = new DataGridViewTextBoxColumn();
            comment_b = new DataGridViewTextBoxColumn();
            remark_b = new DataGridViewTextBoxColumn();
            tabPage3 = new TabPage();
            dataGridView3 = new DataGridView();
            command_c = new DataGridViewTextBoxColumn();
            comment_c = new DataGridViewTextBoxColumn();
            remark_c = new DataGridViewTextBoxColumn();
            open_button = new Button();
            insert_button = new Button();
            clear_button = new Button();
            clear_row_button = new Button();
            bindingSource1 = new BindingSource(components);
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)bindingSource1).BeginInit();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            tabControl1.HotTrack = true;
            tabControl1.ItemSize = new Size(42, 21);
            tabControl1.Location = new Point(12, 12);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(771, 533);
            tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(dataGridView1);
            tabPage1.Location = new Point(4, 25);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(763, 504);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "SatA";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView1.BackgroundColor = SystemColors.AppWorkspace;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.Control;
            dataGridViewCellStyle1.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle1.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.False;
            dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { command_a, comment_a, remark_a });
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = SystemColors.Window;
            dataGridViewCellStyle2.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle2.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            dataGridView1.GridColor = SystemColors.ControlDark;
            dataGridView1.Location = new Point(5, 5);
            dataGridView1.Margin = new Padding(2);
            dataGridView1.MultiSelect = false;
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.Size = new Size(753, 496);
            dataGridView1.TabIndex = 0;
            dataGridView1.CellContentDoubleClick += dataGridView1_CellContentDoubleClick;
            // 
            // command_a
            // 
            command_a.HeaderText = "Command";
            command_a.MinimumWidth = 6;
            command_a.Name = "command_a";
            command_a.Width = 260;
            // 
            // comment_a
            // 
            comment_a.HeaderText = "Comment";
            comment_a.MinimumWidth = 6;
            comment_a.Name = "comment_a";
            comment_a.Width = 260;
            // 
            // remark_a
            // 
            remark_a.HeaderText = "Remark";
            remark_a.MinimumWidth = 6;
            remark_a.Name = "remark_a";
            remark_a.Width = 180;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(dataGridView2);
            tabPage2.Location = new Point(4, 25);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(763, 504);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "SatB";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // dataGridView2
            // 
            dataGridView2.AllowUserToOrderColumns = true;
            dataGridView2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView2.BackgroundColor = SystemColors.AppWorkspace;
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = SystemColors.Control;
            dataGridViewCellStyle3.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle3.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
            dataGridView2.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle3;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Columns.AddRange(new DataGridViewColumn[] { command_B, comment_b, remark_b });
            dataGridViewCellStyle4.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = SystemColors.Window;
            dataGridViewCellStyle4.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            dataGridViewCellStyle4.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = DataGridViewTriState.False;
            dataGridView2.DefaultCellStyle = dataGridViewCellStyle4;
            dataGridView2.GridColor = SystemColors.ControlDark;
            dataGridView2.Location = new Point(5, 5);
            dataGridView2.Margin = new Padding(2);
            dataGridView2.MultiSelect = false;
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowHeadersWidth = 51;
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView2.Size = new Size(753, 296);
            dataGridView2.TabIndex = 1;
            dataGridView2.CellContentDoubleClick += dataGridView2_CellContentDoubleClick;
            // 
            // command_B
            // 
            command_B.HeaderText = "Command";
            command_B.MinimumWidth = 6;
            command_B.Name = "command_B";
            command_B.Width = 260;
            // 
            // comment_b
            // 
            comment_b.HeaderText = "Comment";
            comment_b.MinimumWidth = 6;
            comment_b.Name = "comment_b";
            comment_b.Width = 260;
            // 
            // remark_b
            // 
            remark_b.HeaderText = "Remark";
            remark_b.MinimumWidth = 6;
            remark_b.Name = "remark_b";
            remark_b.Width = 180;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(dataGridView3);
            tabPage3.Location = new Point(4, 25);
            tabPage3.Name = "tabPage3";
            tabPage3.Size = new Size(763, 504);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Another Command";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // dataGridView3
            // 
            dataGridView3.AllowUserToOrderColumns = true;
            dataGridView3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView3.BackgroundColor = SystemColors.AppWorkspace;
            dataGridView3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView3.Columns.AddRange(new DataGridViewColumn[] { command_c, comment_c, remark_c });
            dataGridView3.GridColor = SystemColors.ControlDark;
            dataGridView3.Location = new Point(5, 5);
            dataGridView3.Margin = new Padding(2);
            dataGridView3.MultiSelect = false;
            dataGridView3.Name = "dataGridView3";
            dataGridView3.RowHeadersWidth = 51;
            dataGridView3.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView3.Size = new Size(753, 296);
            dataGridView3.TabIndex = 2;
            dataGridView3.CellContentDoubleClick += dataGridView3_CellContentDoubleClick;
            // 
            // command_c
            // 
            command_c.HeaderText = "Command";
            command_c.MinimumWidth = 6;
            command_c.Name = "command_c";
            command_c.Width = 260;
            // 
            // comment_c
            // 
            comment_c.HeaderText = "Comment";
            comment_c.MinimumWidth = 6;
            comment_c.Name = "comment_c";
            comment_c.Width = 260;
            // 
            // remark_c
            // 
            remark_c.HeaderText = "Remark";
            remark_c.MinimumWidth = 6;
            remark_c.Name = "remark_c";
            remark_c.Width = 180;
            // 
            // open_button
            // 
            open_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            open_button.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            open_button.Location = new Point(14, 546);
            open_button.Margin = new Padding(2);
            open_button.Name = "open_button";
            open_button.Size = new Size(91, 41);
            open_button.TabIndex = 1;
            open_button.Text = "Open";
            open_button.UseVisualStyleBackColor = true;
            open_button.Click += open_button_Click;
            // 
            // insert_button
            // 
            insert_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            insert_button.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            insert_button.Location = new Point(689, 546);
            insert_button.Margin = new Padding(3, 2, 3, 2);
            insert_button.Name = "insert_button";
            insert_button.Size = new Size(91, 41);
            insert_button.TabIndex = 3;
            insert_button.Text = "Insert";
            insert_button.UseVisualStyleBackColor = true;
            insert_button.Click += insert_button_Click;
            // 
            // clear_button
            // 
            clear_button.Anchor = AnchorStyles.Bottom;
            clear_button.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            clear_button.Location = new Point(206, 546);
            clear_button.Margin = new Padding(3, 2, 3, 2);
            clear_button.Name = "clear_button";
            clear_button.Size = new Size(116, 41);
            clear_button.TabIndex = 4;
            clear_button.Text = "Clear All";
            clear_button.UseVisualStyleBackColor = true;
            clear_button.Click += clear_button_Click;
            // 
            // clear_row_button
            // 
            clear_row_button.Anchor = AnchorStyles.Bottom;
            clear_row_button.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            clear_row_button.Location = new Point(449, 546);
            clear_row_button.Margin = new Padding(3, 2, 3, 2);
            clear_row_button.Name = "clear_row_button";
            clear_row_button.Size = new Size(128, 41);
            clear_row_button.TabIndex = 5;
            clear_row_button.Text = "Clear Row";
            clear_row_button.UseVisualStyleBackColor = true;
            clear_row_button.Click += clear_row_button_Click;
            // 
            // bindingSource1
            // 
            bindingSource1.CurrentChanged += bindingSource1_CurrentChanged;
            // 
            // commandList
            // 
            AutoScaleDimensions = new SizeF(9F, 18F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(787, 593);
            ControlBox = false;
            Controls.Add(clear_row_button);
            Controls.Add(clear_button);
            Controls.Add(insert_button);
            Controls.Add(open_button);
            Controls.Add(tabControl1);
            Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "commandList";
            StartPosition = FormStartPosition.Manual;
            Text = "Command List";
            Load += CommandList_Load;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridView3).EndInit();
            ((System.ComponentModel.ISupportInitialize)bindingSource1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage1;
        private DataGridView dataGridView1;
        private TabPage tabPage2;
        private TabPage tabPage3;
        private DataGridView dataGridView2;
        private DataGridViewTextBoxColumn command_a;
        private DataGridViewTextBoxColumn comment_a;
        private DataGridViewTextBoxColumn remark_a;
        private DataGridViewTextBoxColumn command_B;
        private DataGridViewTextBoxColumn comment_b;
        private DataGridViewTextBoxColumn remark_b;
        private DataGridView dataGridView3;
        private DataGridViewTextBoxColumn command_c;
        private DataGridViewTextBoxColumn comment_c;
        private DataGridViewTextBoxColumn remark_c;
        private Button open_button;
        private Button insert_button;
        private Button clear_button;
        private Button clear_row_button;
        private BindingSource bindingSource1;
    }
}