using System.Runtime.Serialization;
using System.Text.RegularExpressions;

using System.IO.Ports;
using System.IO;
using System.Xml.Linq;
using Microsoft.VisualBasic.FileIO;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace S2S_Ground_Station
{
    public partial class operationMain : Form
    {

        public commandList commandList;

        public string CommandString
        {
            set
            {
                command_HEX.Text = value;
            }
            get
            {
                return command_HEX.Text;
            }
        }

        public string CommentString
        {
            set
            {
                comments.Text = value;
            }
            get
            {
                return comments.Text;
            }
        }

        private SerialPort RadioPort = new SerialPort();

        public string call_sign;

        public string savingRootPath;
        public string rootPath;
        public DateTime dtAppStart = DateTime.UtcNow;

        public string freqPath;

        public double latitude, longitude, altitude;
        public double azHomeValue, elHomeValue;

        private bool RadioPortConnected = false;

        private bool freqSelected = false;

        List<string> upFreqList = new List<string>();
        public string upFreq;

        private delegate void AddRxDataDelegate(string data);

        int receivedPacketNumber = 0;

        private bool ReceivedDataFrameFlag = false;

        public operationMain()
        {
            InitializeComponent();
            RadioPort.DataReceived += new SerialDataReceivedEventHandler(DataReceived);
            commandList = new commandList();
            commandList.formMain = this;
            commandList.Show();

            updateSavingFolderPath();
        }
        private void setSatelliteComboList()
        {
            //Parse tle from file
            if (!System.IO.File.Exists(freqPath))
            {
                MessageBox.Show("Frequency Fiile doesn't exist!");
                return;
            }

            satelliteNameComboBox.DataSource = null;
            satelliteNameComboBox.Items.Clear();

            using (var parser = new TextFieldParser(freqPath))
            {
                if (parser != null)
                {

                    parser.Delimiters = new string[] { "," };

                    // Skip the header row
                    parser.ReadLine();

                    while (!parser.EndOfData)
                    {
                        var fields = parser.ReadFields();

                        if (fields != null)
                        {
                            // Add only the first column (fields[0]) to the ComboBox
                            satelliteNameComboBox.Items.Add(fields[0]);

                            // Check if the value in fields[3] is a valid number
                            if (double.TryParse(fields[3], out double frequencyValue))
                            {
                                // Divide the value by 1,000,000 and assign to the label
                                upFreq = (frequencyValue / 1000000.0).ToString("F6");
                                upFreqList.Add(upFreq);
                            }
                        }
                    }
                }
            }
        }

        private void updateSavingFolderPath()
        {
            rootPath = Properties.Settings.Default.saveFilePath;

            savingRootPath = System.IO.Path.Combine(rootPath, dtAppStart.ToString("yyyyMMdd"));

            if (!Directory.Exists(rootPath))
            {
                Directory.CreateDirectory(rootPath);
            }
            if (!Directory.Exists(savingRootPath))
            {
                Directory.CreateDirectory(savingRootPath);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // disable transmit button
            transmitButton.Enabled = false;

            // when program is loaded the groundstation is in telemetry mode
            radio_button_telemetry.Checked = false;
            radio_button_digipeater.Checked = false;

            radio_button_telemetry.Enabled = false;
            radio_button_digipeater.Enabled = false;

            // when program is loaded the com port are disabled
            radioPortComboBox.Enabled = false;
            radioBaudComboBox.Enabled = false;
            radioButtonConnect.Enabled = false;

            //reset the path variable
            savingFolderPathTextBox.Text = null;
            savingFolderPathTextBox.Text = savingRootPath;

            //disable reveived data and save
            //  saveReceiveButton.Enabled = false;

            //Add port list to each COM port ComboBox
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                radioPortComboBox.Items.Add(port);
            }

            //enabling all
            radioPortComboBox.Enabled = true;
            radioBaudComboBox.Enabled = true;
            radioButtonConnect.Enabled = true;

            if (radioPortComboBox.Items.Count >= 1)
            {
                radioPortComboBox.SelectedIndex = 0;
            }

            freqPathTextBox.Text = null;
            //        freqPath = Properties.Settings.Default.freqPath;
            //  freqPathTextBox.Text = freqPath;

            latitudeTextBox.Text = null; longitudeTextBox.Text = null; altitudeTextBox.Text = null;

            latitudeTextBox.Text = Properties.Settings.Default.latitude.ToString("0.0000");
            longitudeTextBox.Text = Properties.Settings.Default.longitude.ToString("0.0000");
            altitudeTextBox.Text = Properties.Settings.Default.altitude.ToString("0.0");

            latitude = Convert.ToDouble(latitudeTextBox.Text);
            longitude = Convert.ToDouble(longitudeTextBox.Text);
            altitude = Convert.ToDouble(altitudeTextBox.Text);

            azHomeValue = Properties.Settings.Default.azHome;
            elHomeValue = Properties.Settings.Default.elHome;

            azimuthLabel.Text = azHomeValue.ToString(".0");
            elevationLabel.Text = elHomeValue.ToString(".0");

            call_sign = Properties.Settings.Default.callsign;
            callsignLabel.Text = call_sign;
        }

        private void transmitButton_Click(object sender, EventArgs e)
        {
            transmitAsync();
        }

        public async void transmitAsync()
        {

            // Disable the button
            transmitButton.Enabled = false;

            // Run a task with a delay (e.g., 2000 ms delay)
            await Task.Run(async () =>
            {
                // Wait for (777 milliseconds) prevent multiple tx at a time
                await Task.Delay(777);
            });

            string command = command_HEX.Text;
            string comment = comments.Text;

            var baudrate = radioPortComboBox.Text;

            if (radio_button_telemetry.Checked == true)
            {

                if (Regex.IsMatch(command, ".. .. .. .. .. .. .. .. .. .. .. .. .."))
                {
                    Console.WriteLine("13 bytes command");
                    command = command.Substring(0, 38);
                }
                else
                {
                    Console.WriteLine("Invalid Length");
                    MessageBox.Show("Invalid Length");
                    return;
                }

                if (command == "00 00 00 00 00 00 00 00 00 00 00 00 00")
                {
                    DialogResult result = MessageBox.Show(
                        "The command is all '00'\n",
                        "Invalid Command",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Exclamation,
                        MessageBoxDefaultButton.Button2);
                    Console.WriteLine("Dialog");
                    if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                else
                {

                    if (RadioPort.IsOpen)
                    {

                        RadioPort.WriteLine(command);

                        DateTime dt = DateTime.Now;
                        string time = dt.ToString("yyyy/MM/dd HH:mm:ss");

                        this.rxDataTextBox.Focus();
                        this.rxDataTextBox.AppendText("#(" + time + ")" + " CMD: " + command + "\n\n");
                    }
                }
                // Enable the button after the delay
                transmitButton.Invoke((Action)(() => transmitButton.Enabled = true));
            }
            else //digipeater is Checked
            {

            }
        }

        private void ClearRxDataButton_Click(object sender, EventArgs e)
        {
            receivedPacketNumber = 0;
            rxPacketCount.Text = receivedPacketNumber.ToString();
            rxDataTextBox.Text = "";
        }

        private void saveReceiveButton_Click(object sender, EventArgs e)
        {
            ClearRxDataButton_Click(sender, e);
        }

        private void command_HEX_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void radioButtonConnect_Click(object sender, EventArgs e)
        {
            try
            {
                radio_button_telemetry.Checked = true;

                if (!RadioPortConnected)
                {
                    RadioPort.PortName = radioPortComboBox.Text;
                    RadioPort.BaudRate = Convert.ToInt32(radioBaudComboBox.SelectedItem);
                    RadioPort.DataBits = 8;
                    RadioPort.Parity = Parity.None;
                    RadioPort.StopBits = StopBits.One;
                    RadioPort.RtsEnable = true;
                    RadioPort.Open();
                    RadioPortConnected = true;
                    radioPortComboBox.Enabled = false;
                    radio_button_telemetry.Enabled = true;
                    radio_button_digipeater.Enabled = true;
                    radioBaudComboBox.Enabled = false;
                    radioButtonConnect.Text = "Disconnect";
                }
                else
                {
                    RadioPort.Close();
                    RadioPortConnected = false;
                    radioPortComboBox.Enabled = true;
                    radio_button_telemetry.Enabled = false;
                    radio_button_digipeater.Enabled = false;
                    radioBaudComboBox.Enabled = true;
                    radioButtonConnect.Text = "Connect";
                }
            }
            catch (Exception ex)
            {
                DialogResult result = MessageBox.Show(
                               $"Problem in connection to {radioPortComboBox.Text}\n {ex.Message}",
                               "Port not found",
                               MessageBoxButtons.OKCancel,
                               MessageBoxIcon.Exclamation,
                               MessageBoxDefaultButton.Button2);
                Console.WriteLine("Dialog");
                if (result == DialogResult.Cancel)
                {
                    return;
                }
            }
            finally
            {
                if (RadioPortConnected && freqSelected)
                {
                    transmitButton.Enabled = true;
                }
                else
                {
                    transmitButton.Enabled = false;
                }
            }
        }

        private void refreshButton_Click(object sender, EventArgs e)
        {
            radioPortComboBox.DataSource = null; // Unbind the data source
            radioPortComboBox.Items.Clear();

            if (!RadioPortConnected)
            {
                radioPortComboBox.Items.Clear();
                //add port list into each COM Port ComboBox
                string[] ports = SerialPort.GetPortNames();
                radioPortComboBox.Items.Clear();

                foreach (string port in ports)
                {
                    radioPortComboBox.Items.Add(port);
                }

                //Number of COM port is more than 1 -> Radio COM port is enabled
                if (radioPortComboBox.Items.Count >= 1)
                {
                    radioPortComboBox.SelectedIndex = 0;
                }
            }
            else
            {
                MessageBox.Show("Please disconnect all com port");
            }
        }

        private void latitudeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (latitudeTextBox.Text == "")
            {
                return;
            }
            try
            {
                latitude = Convert.ToDouble(latitudeTextBox.Text);
                Properties.Settings.Default.latitude = latitude;
                Properties.Settings.Default.Save();
            }
            catch
            {
                latitudeTextBox.Text = latitude.ToString("0.0000");
            }
        }

        private void longitudeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (longitudeTextBox.Text == "")
            {
                return;
            }
            try
            {
                longitude = Convert.ToDouble(longitudeTextBox.Text);
                Properties.Settings.Default.longitude = longitude;
                Properties.Settings.Default.Save();
            }
            catch
            {
                longitudeTextBox.Text = longitude.ToString("0.0000");
            }
        }

        private void altitudeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (altitudeTextBox.Text == "")
            {
                return;
            }
            try
            {
                altitude = Convert.ToDouble(altitudeTextBox.Text);
                Properties.Settings.Default.altitude = Convert.ToDouble(altitudeTextBox.Text);
                Properties.Settings.Default.Save();
            }
            catch
            {
                altitudeTextBox.Text = altitude.ToString("0.0");
            }
        }

        private void selectSavingFolderPathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "Select saving folder";
            fbd.SelectedPath = savingFolderPathTextBox.Text;
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                Console.WriteLine(fbd.SelectedPath);
                savingRootPath = fbd.SelectedPath;
                savingFolderPathTextBox.Text = savingRootPath;
                Properties.Settings.Default.saveFilePath = savingRootPath;
                Properties.Settings.Default.Save();
            }
        }


        private void selecteFreqPathButton_Click(object sender, EventArgs e)
        {
            //OpenFileDialog class instance
            OpenFileDialog ofd = new OpenFileDialog();

            //specify the first file name
            //specify the string display in the file name
            ofd.FileName = "FreqList.csv";
            //specify the choices that appears in the File Type
            //if specified all files will be displayed
            ofd.Filter = "CSV file(*.csv)|*.*";
            //set the title
            ofd.Title = "Select the frequency list file";
            //the current directory before closing the dialog box
            ofd.RestoreDirectory = true;

            //display dialog
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //dipslay the selected file name when OK button is clicked
                Console.WriteLine(ofd.FileName);
                freqPath = ofd.FileName;
                freqPathTextBox.Text = freqPath;
                Properties.Settings.Default.freqPath = freqPath;
                Properties.Settings.Default.Save();
                setSatelliteComboList();
            }
        }

        private void satelliteNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Get the selected index in the ComboBox
            int selectedIndex = satelliteNameComboBox.SelectedIndex;

            // Check if the selected index is valid (not -1)
            if (selectedIndex != -1)
            {
                // Select the corresponding index in the uList (the data is already linked)
                string selectedValueFromList = upFreqList[selectedIndex];

                freqLabel.Text = selectedValueFromList;
                freqSelected = true;
            }

            if (RadioPortConnected && freqSelected)
            {
                transmitButton.Enabled = true;
            }
            else
            {
                transmitButton.Enabled = false;
            }
        }

        private void AddRxData(string data)
        {
            this.rxDataTextBox.Focus();
            this.rxDataTextBox.AppendText(data);
            receivedPacketNumber += (data.Split('\n').Length - 1);
            rxPacketCount.Text = receivedPacketNumber.ToString();
        }

        private void DataReceived(object Sender, SerialDataReceivedEventArgs e)
        {
            AddRxDataDelegate add = new AddRxDataDelegate(AddRxData);
            byte[] SerialGetData = new byte[RadioPort.BytesToRead];
            string ReceivedDataString = string.Empty;

            try
            {
                RadioPort.Read(SerialGetData, 0, SerialGetData.Length);
            }

            catch
            {
                //Protect error
            }


            for (int i = 0; i < SerialGetData.Length; i++)
            {
                byte EachData = SerialGetData[i];
                ReceivedDataString += string.Format("{0:x2} ", EachData);

                if (EachData == 0x7e)
                {
                    if (ReceivedDataFrameFlag)
                    {
                        ReceivedDataFrameFlag = false;
                        ReceivedDataString += "\n\n";
                    }
                    else
                    {
                        ReceivedDataFrameFlag = true;
                    }
                }
            }

            DateTime dt = DateTime.Now;
            string time = dt.ToString("yyyy/MM/dd HH:mm:ss");

            string ptype = string.Empty;
            try
            {
                if (SerialGetData[17] == 172)    //ac
                {
                    ptype = " ACK";
                }
                else if (SerialGetData[17] == 177)  //b1
                {
                    ptype = " BA_1";
                }
                else if (SerialGetData[17] == 178) //b2
                {
                    ptype = " BA_2";
                }
                else
                {
                    ptype = " RES";
                }
            }
            catch (Exception ex)
            {
                ptype = " RES";
            }
            finally
            {
                // this.rxDataTextBox.Invoke(add, "#(" + time + ")" + ptype + ": " + ReceivedDataString);
                this.rxDataTextBox.Invoke(add, ReceivedDataString);
                ReceivedDataString = string.Empty;
                SerialGetData = new byte[0];
            }
        }
    }
}
