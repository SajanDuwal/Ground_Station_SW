using System.Runtime.CompilerServices;

namespace S2S_Ground_Station
{
    partial class operationMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(operationMain));
            pictureBox1 = new PictureBox();
            main_tab_control = new TabControl();
            tabPage1 = new TabPage();
            rxPacketCount = new Label();
            label8 = new Label();
            comments = new Label();
            ReceivedLabel = new Label();
            label5 = new Label();
            label3 = new Label();
            groupBox1 = new GroupBox();
            radio_button_digipeater = new RadioButton();
            radio_button_telemetry = new RadioButton();
            rxDataTextBox = new RichTextBox();
            command_HEX = new MaskedTextBox();
            saveReceiveButton = new Button();
            transmitButton = new Button();
            satelliteNameComboBox = new ComboBox();
            label7 = new Label();
            tabPage2 = new TabPage();
            groupBox4 = new GroupBox();
            altitudeTextBox = new TextBox();
            longitudeTextBox = new TextBox();
            latitudeTextBox = new TextBox();
            label14 = new Label();
            label13 = new Label();
            label12 = new Label();
            groupBox3 = new GroupBox();
            selectSavingFolderPathButton = new Button();
            savingFolderPathTextBox = new TextBox();
            selecteFreqPathButton = new Button();
            freqPathTextBox = new TextBox();
            label11 = new Label();
            label10 = new Label();
            groupBox2 = new GroupBox();
            refreshButton = new Button();
            radioButtonConnect = new Button();
            radioBaudComboBox = new ComboBox();
            radioPortComboBox = new ComboBox();
            label9 = new Label();
            label1 = new Label();
            azimuthLabel = new Label();
            elevationLabel = new Label();
            label4 = new Label();
            freqLabel = new Label();
            label6 = new Label();
            pictureBox2 = new PictureBox();
            pictureBox3 = new PictureBox();
            callsignLabel = new Label();
            label2 = new Label();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            main_tab_control.SuspendLayout();
            tabPage1.SuspendLayout();
            groupBox1.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).BeginInit();
            SuspendLayout();
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pictureBox1.ErrorImage = (Image)resources.GetObject("pictureBox1.ErrorImage");
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.InitialImage = (Image)resources.GetObject("pictureBox1.InitialImage");
            pictureBox1.Location = new Point(897, 8);
            pictureBox1.Margin = new Padding(3, 2, 3, 2);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(124, 113);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.TabIndex = 0;
            pictureBox1.TabStop = false;
            // 
            // main_tab_control
            // 
            main_tab_control.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            main_tab_control.Controls.Add(tabPage1);
            main_tab_control.Controls.Add(tabPage2);
            main_tab_control.Font = new Font("Microsoft Sans Serif", 7.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            main_tab_control.ItemSize = new Size(48, 21);
            main_tab_control.Location = new Point(17, 123);
            main_tab_control.Name = "main_tab_control";
            main_tab_control.SelectedIndex = 0;
            main_tab_control.Size = new Size(1008, 820);
            main_tab_control.TabIndex = 46;
            // 
            // tabPage1
            // 
            tabPage1.BackColor = SystemColors.Window;
            tabPage1.Controls.Add(rxPacketCount);
            tabPage1.Controls.Add(label8);
            tabPage1.Controls.Add(comments);
            tabPage1.Controls.Add(ReceivedLabel);
            tabPage1.Controls.Add(label5);
            tabPage1.Controls.Add(label3);
            tabPage1.Controls.Add(groupBox1);
            tabPage1.Controls.Add(rxDataTextBox);
            tabPage1.Controls.Add(command_HEX);
            tabPage1.Controls.Add(saveReceiveButton);
            tabPage1.Controls.Add(transmitButton);
            tabPage1.Controls.Add(satelliteNameComboBox);
            tabPage1.Controls.Add(label7);
            tabPage1.Location = new Point(4, 25);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(1000, 791);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Main";
            // 
            // rxPacketCount
            // 
            rxPacketCount.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            rxPacketCount.AutoSize = true;
            rxPacketCount.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            rxPacketCount.Location = new Point(235, 733);
            rxPacketCount.Name = "rxPacketCount";
            rxPacketCount.Size = new Size(22, 23);
            rxPacketCount.TabIndex = 72;
            rxPacketCount.Text = "0";
            // 
            // label8
            // 
            label8.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label8.AutoSize = true;
            label8.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label8.Location = new Point(11, 732);
            label8.Name = "label8";
            label8.Size = new Size(218, 23);
            label8.TabIndex = 71;
            label8.Text = "Received Packet(s) : ";
            // 
            // comments
            // 
            comments.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            comments.AutoEllipsis = true;
            comments.BackColor = Color.Gainsboro;
            comments.BorderStyle = BorderStyle.Fixed3D;
            comments.Font = new Font("MS UI Gothic", 13.2000008F, FontStyle.Regular, GraphicsUnit.Point, 0);
            comments.Location = new Point(188, 150);
            comments.Margin = new Padding(0);
            comments.Name = "comments";
            comments.Padding = new Padding(3, 2, 3, 2);
            comments.Size = new Size(620, 70);
            comments.TabIndex = 5;
            comments.Text = "None";
            // 
            // ReceivedLabel
            // 
            ReceivedLabel.AutoSize = true;
            ReceivedLabel.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            ReceivedLabel.Location = new Point(11, 232);
            ReceivedLabel.Name = "ReceivedLabel";
            ReceivedLabel.Size = new Size(246, 23);
            ReceivedLabel.TabIndex = 70;
            ReceivedLabel.Text = "CMD and Received Data";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label5.Location = new Point(11, 150);
            label5.Name = "label5";
            label5.Size = new Size(114, 23);
            label5.TabIndex = 69;
            label5.Text = "Comments";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label3.Location = new Point(11, 110);
            label3.Name = "label3";
            label3.Size = new Size(168, 23);
            label3.TabIndex = 68;
            label3.Text = "Command (HEX)";
            // 
            // groupBox1
            // 
            groupBox1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            groupBox1.Controls.Add(radio_button_digipeater);
            groupBox1.Controls.Add(radio_button_telemetry);
            groupBox1.Location = new Point(615, 18);
            groupBox1.Margin = new Padding(3, 2, 3, 2);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(354, 77);
            groupBox1.TabIndex = 13;
            groupBox1.TabStop = false;
            groupBox1.Text = "Radio Mode";
            // 
            // radio_button_digipeater
            // 
            radio_button_digipeater.AutoSize = true;
            radio_button_digipeater.Location = new Point(215, 30);
            radio_button_digipeater.Name = "radio_button_digipeater";
            radio_button_digipeater.Size = new Size(91, 20);
            radio_button_digipeater.TabIndex = 1;
            radio_button_digipeater.TabStop = true;
            radio_button_digipeater.Text = "Digipeater";
            radio_button_digipeater.UseVisualStyleBackColor = true;
            // 
            // radio_button_telemetry
            // 
            radio_button_telemetry.AutoSize = true;
            radio_button_telemetry.Location = new Point(27, 30);
            radio_button_telemetry.Name = "radio_button_telemetry";
            radio_button_telemetry.Size = new Size(89, 20);
            radio_button_telemetry.TabIndex = 0;
            radio_button_telemetry.TabStop = true;
            radio_button_telemetry.Text = "Telemetry";
            radio_button_telemetry.UseVisualStyleBackColor = true;
            // 
            // rxDataTextBox
            // 
            rxDataTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            rxDataTextBox.Cursor = Cursors.IBeam;
            rxDataTextBox.Font = new Font("MS Gothic", 13.2000008F, FontStyle.Regular, GraphicsUnit.Point, 0);
            rxDataTextBox.ForeColor = Color.Black;
            rxDataTextBox.Location = new Point(11, 260);
            rxDataTextBox.Margin = new Padding(3, 2, 3, 2);
            rxDataTextBox.Name = "rxDataTextBox";
            rxDataTextBox.ScrollBars = RichTextBoxScrollBars.ForcedBoth;
            rxDataTextBox.Size = new Size(973, 439);
            rxDataTextBox.TabIndex = 14;
            rxDataTextBox.Text = "";
            // 
            // command_HEX
            // 
            command_HEX.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            command_HEX.Cursor = Cursors.IBeam;
            command_HEX.Font = new Font("MS Gothic", 13.2000008F, FontStyle.Regular, GraphicsUnit.Point, 0);
            command_HEX.Location = new Point(188, 110);
            command_HEX.Margin = new Padding(3, 2, 3, 2);
            command_HEX.Mask = ">AA AA AA AA AA AA AA AA AA AA AA AA AA";
            command_HEX.Name = "command_HEX";
            command_HEX.Size = new Size(776, 29);
            command_HEX.TabIndex = 2;
            command_HEX.Text = "00000000000000000000000000";
            command_HEX.MaskInputRejected += command_HEX_MaskInputRejected;
            // 
            // saveReceiveButton
            // 
            saveReceiveButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            saveReceiveButton.BackColor = Color.Transparent;
            saveReceiveButton.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            saveReceiveButton.Location = new Point(817, 712);
            saveReceiveButton.Margin = new Padding(3, 2, 3, 2);
            saveReceiveButton.Name = "saveReceiveButton";
            saveReceiveButton.Size = new Size(177, 59);
            saveReceiveButton.TabIndex = 15;
            saveReceiveButton.Text = "Save and Clear";
            saveReceiveButton.UseVisualStyleBackColor = false;
            saveReceiveButton.Click += saveReceiveButton_Click;
            // 
            // transmitButton
            // 
            transmitButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            transmitButton.Font = new Font("MS UI Gothic", 15F, FontStyle.Regular, GraphicsUnit.Point, 0);
            transmitButton.Location = new Point(817, 150);
            transmitButton.Margin = new Padding(3, 2, 3, 2);
            transmitButton.Name = "transmitButton";
            transmitButton.Size = new Size(152, 69);
            transmitButton.TabIndex = 1;
            transmitButton.Text = "Transmit";
            transmitButton.UseMnemonic = false;
            transmitButton.UseVisualStyleBackColor = true;
            transmitButton.Click += transmitButton_Click;
            // 
            // satelliteNameComboBox
            // 
            satelliteNameComboBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            satelliteNameComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            satelliteNameComboBox.DropDownWidth = 41;
            satelliteNameComboBox.Font = new Font("MS Gothic", 18F, FontStyle.Regular, GraphicsUnit.Point, 0);
            satelliteNameComboBox.ForeColor = SystemColors.Highlight;
            satelliteNameComboBox.FormattingEnabled = true;
            satelliteNameComboBox.Location = new Point(15, 54);
            satelliteNameComboBox.Margin = new Padding(3, 2, 3, 2);
            satelliteNameComboBox.Name = "satelliteNameComboBox";
            satelliteNameComboBox.Size = new Size(400, 38);
            satelliteNameComboBox.TabIndex = 29;
            satelliteNameComboBox.SelectedIndexChanged += satelliteNameComboBox_SelectedIndexChanged;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("MS UI Gothic", 13.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label7.Location = new Point(11, 18);
            label7.Name = "label7";
            label7.Size = new Size(152, 23);
            label7.TabIndex = 30;
            label7.Text = "Satellite Name";
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(groupBox4);
            tabPage2.Controls.Add(groupBox3);
            tabPage2.Controls.Add(groupBox2);
            tabPage2.Location = new Point(4, 25);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(1000, 791);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Setting";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBox4.Controls.Add(altitudeTextBox);
            groupBox4.Controls.Add(longitudeTextBox);
            groupBox4.Controls.Add(latitudeTextBox);
            groupBox4.Controls.Add(label14);
            groupBox4.Controls.Add(label13);
            groupBox4.Controls.Add(label12);
            groupBox4.Location = new Point(11, 216);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(971, 125);
            groupBox4.TabIndex = 46;
            groupBox4.TabStop = false;
            groupBox4.Text = "Ground Station Location";
            // 
            // altitudeTextBox
            // 
            altitudeTextBox.Location = new Point(151, 84);
            altitudeTextBox.Margin = new Padding(4);
            altitudeTextBox.Name = "altitudeTextBox";
            altitudeTextBox.Size = new Size(125, 22);
            altitudeTextBox.TabIndex = 5;
            altitudeTextBox.TextChanged += altitudeTextBox_TextChanged;
            // 
            // longitudeTextBox
            // 
            longitudeTextBox.Cursor = Cursors.IBeam;
            longitudeTextBox.Location = new Point(151, 56);
            longitudeTextBox.Margin = new Padding(4);
            longitudeTextBox.Name = "longitudeTextBox";
            longitudeTextBox.Size = new Size(125, 22);
            longitudeTextBox.TabIndex = 4;
            longitudeTextBox.TextChanged += longitudeTextBox_TextChanged;
            // 
            // latitudeTextBox
            // 
            latitudeTextBox.Cursor = Cursors.IBeam;
            latitudeTextBox.Location = new Point(151, 28);
            latitudeTextBox.Margin = new Padding(4);
            latitudeTextBox.Name = "latitudeTextBox";
            latitudeTextBox.Size = new Size(125, 22);
            latitudeTextBox.TabIndex = 3;
            latitudeTextBox.TextChanged += latitudeTextBox_TextChanged;
            // 
            // label14
            // 
            label14.AutoSize = true;
            label14.Location = new Point(21, 81);
            label14.Margin = new Padding(4, 0, 4, 0);
            label14.Name = "label14";
            label14.Size = new Size(51, 16);
            label14.TabIndex = 2;
            label14.Text = "Altitude";
            // 
            // label13
            // 
            label13.AutoSize = true;
            label13.Location = new Point(21, 54);
            label13.Margin = new Padding(4, 0, 4, 0);
            label13.Name = "label13";
            label13.Size = new Size(66, 16);
            label13.TabIndex = 1;
            label13.Text = "Longitude";
            // 
            // label12
            // 
            label12.AutoSize = true;
            label12.Location = new Point(21, 28);
            label12.Margin = new Padding(4, 0, 4, 0);
            label12.Name = "label12";
            label12.Size = new Size(54, 16);
            label12.TabIndex = 0;
            label12.Text = "Latitude";
            // 
            // groupBox3
            // 
            groupBox3.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBox3.Controls.Add(selectSavingFolderPathButton);
            groupBox3.Controls.Add(savingFolderPathTextBox);
            groupBox3.Controls.Add(selecteFreqPathButton);
            groupBox3.Controls.Add(freqPathTextBox);
            groupBox3.Controls.Add(label11);
            groupBox3.Controls.Add(label10);
            groupBox3.Location = new Point(11, 103);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(971, 107);
            groupBox3.TabIndex = 45;
            groupBox3.TabStop = false;
            groupBox3.Text = "File IO";
            // 
            // selectSavingFolderPathButton
            // 
            selectSavingFolderPathButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            selectSavingFolderPathButton.Location = new Point(924, 68);
            selectSavingFolderPathButton.Margin = new Padding(4);
            selectSavingFolderPathButton.Name = "selectSavingFolderPathButton";
            selectSavingFolderPathButton.Size = new Size(36, 25);
            selectSavingFolderPathButton.TabIndex = 48;
            selectSavingFolderPathButton.Text = "...";
            selectSavingFolderPathButton.UseVisualStyleBackColor = true;
            selectSavingFolderPathButton.Click += selectSavingFolderPathButton_Click;
            // 
            // savingFolderPathTextBox
            // 
            savingFolderPathTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            savingFolderPathTextBox.BackColor = SystemColors.Control;
            savingFolderPathTextBox.Cursor = Cursors.IBeam;
            savingFolderPathTextBox.Location = new Point(151, 68);
            savingFolderPathTextBox.Margin = new Padding(4);
            savingFolderPathTextBox.Name = "savingFolderPathTextBox";
            savingFolderPathTextBox.Size = new Size(765, 22);
            savingFolderPathTextBox.TabIndex = 47;
            // 
            // selecteFreqPathButton
            // 
            selecteFreqPathButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            selecteFreqPathButton.Location = new Point(923, 36);
            selecteFreqPathButton.Margin = new Padding(4);
            selecteFreqPathButton.Name = "selecteFreqPathButton";
            selecteFreqPathButton.Size = new Size(36, 25);
            selecteFreqPathButton.TabIndex = 46;
            selecteFreqPathButton.Text = "...";
            selecteFreqPathButton.UseVisualStyleBackColor = true;
            selecteFreqPathButton.Click += selecteFreqPathButton_Click;
            // 
            // freqPathTextBox
            // 
            freqPathTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            freqPathTextBox.BackColor = SystemColors.Control;
            freqPathTextBox.Cursor = Cursors.IBeam;
            freqPathTextBox.Location = new Point(151, 36);
            freqPathTextBox.Margin = new Padding(4);
            freqPathTextBox.Name = "freqPathTextBox";
            freqPathTextBox.Size = new Size(765, 22);
            freqPathTextBox.TabIndex = 45;
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Location = new Point(21, 38);
            label11.Margin = new Padding(4, 0, 4, 0);
            label11.Name = "label11";
            label11.Size = new Size(94, 16);
            label11.TabIndex = 44;
            label11.Text = "Frequency List";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(21, 68);
            label10.Margin = new Padding(4, 0, 4, 0);
            label10.Name = "label10";
            label10.Size = new Size(101, 16);
            label10.TabIndex = 38;
            label10.Text = "Save To Folder";
            // 
            // groupBox2
            // 
            groupBox2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBox2.Controls.Add(refreshButton);
            groupBox2.Controls.Add(radioButtonConnect);
            groupBox2.Controls.Add(radioBaudComboBox);
            groupBox2.Controls.Add(radioPortComboBox);
            groupBox2.Controls.Add(label9);
            groupBox2.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            groupBox2.Location = new Point(11, 18);
            groupBox2.Margin = new Padding(4);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new Padding(4);
            groupBox2.Size = new Size(971, 78);
            groupBox2.TabIndex = 44;
            groupBox2.TabStop = false;
            groupBox2.Text = "COM Port";
            // 
            // refreshButton
            // 
            refreshButton.Font = new Font("MS UI Gothic", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            refreshButton.Location = new Point(481, 16);
            refreshButton.Margin = new Padding(4);
            refreshButton.Name = "refreshButton";
            refreshButton.Size = new Size(101, 47);
            refreshButton.TabIndex = 46;
            refreshButton.Text = "Refresh COM Port";
            refreshButton.UseVisualStyleBackColor = true;
            refreshButton.Click += refreshButton_Click;
            // 
            // radioButtonConnect
            // 
            radioButtonConnect.Location = new Point(364, 24);
            radioButtonConnect.Margin = new Padding(3, 2, 3, 2);
            radioButtonConnect.Name = "radioButtonConnect";
            radioButtonConnect.Size = new Size(100, 26);
            radioButtonConnect.TabIndex = 11;
            radioButtonConnect.Text = "Connect";
            radioButtonConnect.UseVisualStyleBackColor = true;
            radioButtonConnect.Click += radioButtonConnect_Click;
            // 
            // radioBaudComboBox
            // 
            radioBaudComboBox.FormattingEnabled = true;
            radioBaudComboBox.ItemHeight = 18;
            radioBaudComboBox.Items.AddRange(new object[] { "9600", "57600", "115200" });
            radioBaudComboBox.Location = new Point(235, 24);
            radioBaudComboBox.Margin = new Padding(3, 2, 3, 2);
            radioBaudComboBox.Name = "radioBaudComboBox";
            radioBaudComboBox.Size = new Size(100, 26);
            radioBaudComboBox.TabIndex = 21;
            radioBaudComboBox.Text = "9600";
            // 
            // radioPortComboBox
            // 
            radioPortComboBox.DropDownWidth = 121;
            radioPortComboBox.FormattingEnabled = true;
            radioPortComboBox.ItemHeight = 18;
            radioPortComboBox.Location = new Point(91, 24);
            radioPortComboBox.Margin = new Padding(3, 2, 3, 2);
            radioPortComboBox.Name = "radioPortComboBox";
            radioPortComboBox.Size = new Size(121, 26);
            radioPortComboBox.TabIndex = 9;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(21, 30);
            label9.Name = "label9";
            label9.Size = new Size(47, 18);
            label9.TabIndex = 8;
            label9.Text = "Radio";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Microsoft Sans Serif", 15F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label1.Location = new Point(21, 38);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(52, 29);
            label1.TabIndex = 49;
            label1.Text = "AZ.";
            // 
            // azimuthLabel
            // 
            azimuthLabel.AutoSize = true;
            azimuthLabel.Font = new Font("Microsoft Sans Serif", 21F, FontStyle.Regular, GraphicsUnit.Point, 0);
            azimuthLabel.Location = new Point(69, 31);
            azimuthLabel.Margin = new Padding(4, 0, 4, 0);
            azimuthLabel.Name = "azimuthLabel";
            azimuthLabel.Size = new Size(103, 39);
            azimuthLabel.TabIndex = 52;
            azimuthLabel.Text = "000.0";
            // 
            // elevationLabel
            // 
            elevationLabel.AutoSize = true;
            elevationLabel.Font = new Font("Microsoft Sans Serif", 21F, FontStyle.Regular, GraphicsUnit.Point, 0);
            elevationLabel.Location = new Point(244, 31);
            elevationLabel.Margin = new Padding(4, 0, 4, 0);
            elevationLabel.Name = "elevationLabel";
            elevationLabel.Size = new Size(84, 39);
            elevationLabel.TabIndex = 53;
            elevationLabel.Text = "00.0";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Microsoft Sans Serif", 15F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label4.Location = new Point(194, 38);
            label4.Margin = new Padding(4, 0, 4, 0);
            label4.Name = "label4";
            label4.Size = new Size(51, 29);
            label4.TabIndex = 50;
            label4.Text = "EL.";
            // 
            // freqLabel
            // 
            freqLabel.AutoSize = true;
            freqLabel.BackColor = Color.Black;
            freqLabel.Font = new Font("Microsoft Sans Serif", 22.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            freqLabel.ForeColor = Color.White;
            freqLabel.Location = new Point(87, 78);
            freqLabel.Margin = new Padding(4, 0, 4, 0);
            freqLabel.Name = "freqLabel";
            freqLabel.Size = new Size(217, 42);
            freqLabel.TabIndex = 48;
            freqLabel.Text = "000.000000";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Microsoft Sans Serif", 15F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label6.Location = new Point(17, 86);
            label6.Margin = new Padding(4, 0, 4, 0);
            label6.Name = "label6";
            label6.Size = new Size(71, 29);
            label6.TabIndex = 52;
            label6.Text = "Freq.";
            // 
            // pictureBox2
            // 
            pictureBox2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pictureBox2.ErrorImage = (Image)resources.GetObject("pictureBox2.ErrorImage");
            pictureBox2.Image = (Image)resources.GetObject("pictureBox2.Image");
            pictureBox2.InitialImage = (Image)resources.GetObject("pictureBox2.InitialImage");
            pictureBox2.Location = new Point(788, 8);
            pictureBox2.Margin = new Padding(3, 2, 3, 2);
            pictureBox2.Name = "pictureBox2";
            pictureBox2.Size = new Size(114, 113);
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox2.TabIndex = 70;
            pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            pictureBox3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pictureBox3.ErrorImage = (Image)resources.GetObject("pictureBox3.ErrorImage");
            pictureBox3.Image = (Image)resources.GetObject("pictureBox3.Image");
            pictureBox3.InitialImage = (Image)resources.GetObject("pictureBox3.InitialImage");
            pictureBox3.Location = new Point(495, 8);
            pictureBox3.Margin = new Padding(3, 2, 3, 2);
            pictureBox3.Name = "pictureBox3";
            pictureBox3.Size = new Size(287, 113);
            pictureBox3.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox3.TabIndex = 71;
            pictureBox3.TabStop = false;
            // 
            // callsignLabel
            // 
            callsignLabel.AutoSize = true;
            callsignLabel.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            callsignLabel.ForeColor = Color.Red;
            callsignLabel.Location = new Point(104, 9);
            callsignLabel.Margin = new Padding(4, 0, 4, 0);
            callsignLabel.Name = "callsignLabel";
            callsignLabel.Size = new Size(102, 20);
            callsignLabel.TabIndex = 72;
            callsignLabel.Text = "_ _ _ _ _ _ _";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Microsoft Sans Serif", 10.2F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label2.Location = new Point(17, 8);
            label2.Margin = new Padding(4, 0, 4, 0);
            label2.Name = "label2";
            label2.Size = new Size(79, 20);
            label2.TabIndex = 73;
            label2.Text = "Callsign :";
            // 
            // operationMain
            // 
            AutoScaleDimensions = new SizeF(8F, 16F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1045, 953);
            Controls.Add(label2);
            Controls.Add(callsignLabel);
            Controls.Add(pictureBox3);
            Controls.Add(pictureBox2);
            Controls.Add(freqLabel);
            Controls.Add(label6);
            Controls.Add(elevationLabel);
            Controls.Add(label4);
            Controls.Add(azimuthLabel);
            Controls.Add(label1);
            Controls.Add(main_tab_control);
            Controls.Add(pictureBox1);
            Font = new Font("Microsoft Sans Serif", 7.8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "operationMain";
            Text = "beta Slipper2Sat Ground Station";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            main_tab_control.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            tabPage2.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)pictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox3).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private PictureBox pictureBox1;
        private TabControl main_tab_control;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Label label1;
        private Label azimuthLabel;
        private Label elevationLabel;
        private Label label4;
        private Label freqLabel;
        private Label label6;
        private ComboBox satelliteNameComboBox;
        private Label label7;
        private Button saveReceiveButton;
        private Button transmitButton;
        private RichTextBox rxDataTextBox;
        private MaskedTextBox command_HEX;
        private GroupBox groupBox1;
        private RadioButton radio_button_digipeater;
        private RadioButton radio_button_telemetry;
        private PictureBox pictureBox2;
        private PictureBox pictureBox3;
        private Label callsignLabel;
        private Label label2;
        private Label label5;
        private Label label3;
        private Label ReceivedLabel;
        private Label comments;
        private Label label8;
        private Label rxPacketCount;
        private GroupBox groupBox2;
        private Label label9;
        private Button refreshButton;
        private Button radioButtonConnect;
        private ComboBox radioBaudComboBox;
        private ComboBox radioPortComboBox;
        private GroupBox groupBox4;
        private GroupBox groupBox3;
        private Label label11;
        private Label label10;
        private TextBox altitudeTextBox;
        private TextBox longitudeTextBox;
        private TextBox latitudeTextBox;
        private Label label14;
        private Label label13;
        private Label label12;
        private Button selecteFreqPathButton;
        private TextBox freqPathTextBox;
        private Button selectSavingFolderPathButton;
        private TextBox savingFolderPathTextBox;
    }
}
