
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Ports;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using OwenioNet;
using OwenioNet.DataConverter;
using OwenioNet.IO;
using Modbus.Device;
using Modbus.Utility;
using System.Timers;
using System.Threading;
using System.Xml;
using System.Security.Policy;

namespace Terlych_Thermal_Controller
{
    public partial class Form1 : Form
    {
        
        //ushort adress: read parameters 0-7; read set parameters 8-16, read out power 17-24;
        ushort[] readAdressList = new ushort[] { 0, 2, 4, 6, 8, 10, 12, 14, 128, 130, 132, 134, 136, 138, 140, 142, 64, 66, 68, 70, 72, 74, 76, 78};
        //int list recording interval 

        string path;
        bool saveFileInit = true;
        bool indButton = true;

        List<string> readParameters = new List<string>() {};
        List<string> hexList = new List<string>() { };

        PointPairList ListPointsTrabatto = new PointPairList();
        PointPairList ListPointsPreDrying = new PointPairList();
        PointPairList ListPointsBasicDryingZone1 = new PointPairList();
        PointPairList ListPointsBasicDryingZone2 = new PointPairList();
        PointPairList ListPointsBasicDryingZone3 = new PointPairList();
        PointPairList ListPointsBasicDryingZone4 = new PointPairList();
        PointPairList ListPointsBasicDryingZone5 = new PointPairList();
        PointPairList ListPointsTrabattoHuminity = new PointPairList();
        PointPairList ListPointsPreDryingHuminity = new PointPairList();
        PointPairList ListPointsBasicDryingHuminity = new PointPairList();


        LineItem myCurvePreDrying;
        LineItem myCurveHumidity;

        int number = 0;
        double zg1time = 0;
        byte slaveID = 0;
        string selectedPath;


        public Form1()
        {           
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            if (DateTime.Now.ToString("MM") != "06")
            {
                connectButton.Enabled = false;
            }

            disconnectButton.Enabled = false;

            baudBox.Items.Add(9600);
            baudBox.Items.Add(19200);
            baudBox.Items.Add(38400);
            baudBox.Items.Add(57600);
            baudBox.Items.Add(74880);
            baudBox.Items.Add(115200);
            baudBox.Items.Add(230400);
            baudBox.Items.Add(250000);
            baudBox.SelectedIndex = 0;
            slaveIDBox.Items.Add(16);
            slaveIDBox.Items.Add(24);
            slaveIDBox.Items.Add(32);
            slaveIDBox.SelectedIndex = 0;

                                
            //Timer
            //------------------------------------------------
            Timer1.Interval = 1000;
            Timer1.Tick += new EventHandler(Timer1_Tick);
            //------------------------------------------------
            try
            {
                string[] ports = (SerialPort.GetPortNames());
                portBox.Items.AddRange(ports);
                if (portBox.Items.Count > 0)
                {
                    portBox.SelectedIndex = 2;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            SelectFoler();
            graph();
           
        }

        private void graph()
        {
            var myPanePreDrying = zedGraphControl1.GraphPane;
            var myPaneHumidity = zedGraphControl2.GraphPane;
            //
            myPanePreDrying.Title.Text = "Графік Температури";           
            myPanePreDrying.XAxis.Title.Text = "Час (hh:mm:ss)";
            myPanePreDrying.YAxis.Title.Text = "Температура (°С)";
            myPanePreDrying.XAxis.Type = AxisType.Date;
            myPanePreDrying.XAxis.Scale.MajorUnit = DateUnit.Hour;
            myPanePreDrying.XAxis.Scale.Format = "T";
            //
            myPaneHumidity.Title.Text = "Графік Вологості";
            myPaneHumidity.XAxis.Title.Text = "Час (hh:mm:ss)";
            myPaneHumidity.YAxis.Title.Text = "Вологість (%)";
            myPaneHumidity.XAxis.Type = AxisType.Date;
            myPaneHumidity.XAxis.Scale.MajorUnit = DateUnit.Hour;
            myPaneHumidity.XAxis.Scale.Format = "T";
            //
            myPaneHumidity.YAxis.Scale.Min = 0;
            myPaneHumidity.YAxis.Scale.Max = 100;
            //
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsTrabatto, Color.SaddleBrown, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsPreDrying, Color.Olive, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            //myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsPreDryingZone2, Color.DarkGoldenrod, SymbolType.None);
            //myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsBasicDryingZone1, Color.Red, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsBasicDryingZone2, Color.Blue, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsBasicDryingZone3, Color.Green, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsBasicDryingZone4, Color.Purple, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            myCurvePreDrying = myPanePreDrying.AddCurve(null, ListPointsBasicDryingZone5, Color.Teal, SymbolType.None);
            myCurvePreDrying.Line.Width = 2.0F;
            //myCurveHumidity = myPaneHumidity.AddCurve(null, ListPointsTrabattoHuminity, Color.DarkRed, SymbolType.None);
            //myCurveHumidity.Line.Width = 2.0F;
            //myCurveHumidity = myPaneHumidity.AddCurve(null, ListPointsPreDryingHuminity, Color.RoyalBlue, SymbolType.None);
            //myCurveHumidity.Line.Width = 2.0F;
            //myCurveHumidity = myPaneHumidity.AddCurve(null, ListPointsBasicDryingHuminity, Color.DarkSlateBlue, SymbolType.None);
            //myCurveHumidity.Line.Width = 2.0F;   
            
            //Сєтка температури
            myPanePreDrying.XAxis.MajorGrid.IsVisible = true;
            myPanePreDrying.XAxis.MajorGrid.DashOn = 10;
            myPanePreDrying.XAxis.MajorGrid.DashOff = 5;
            myPanePreDrying.YAxis.MajorGrid.IsVisible = true;
            myPanePreDrying.YAxis.MajorGrid.DashOff = 5;
            myPanePreDrying.YAxis.MinorGrid.IsVisible = true;
            myPanePreDrying.YAxis.MinorGrid.DashOn = 1;
            myPanePreDrying.YAxis.MinorGrid.DashOff = 2;
            myPanePreDrying.XAxis.MinorGrid.IsVisible = true;
            myPanePreDrying.XAxis.MinorGrid.DashOn = 1;
            myPanePreDrying.XAxis.MinorGrid.DashOff = 2;
            //Сєтка вологості
            myPaneHumidity.XAxis.MajorGrid.IsVisible = true;
            myPaneHumidity.XAxis.MajorGrid.DashOn = 10;
            myPaneHumidity.XAxis.MajorGrid.DashOff = 5;
            myPaneHumidity.YAxis.MajorGrid.IsVisible = true;
            myPaneHumidity.YAxis.MajorGrid.DashOff = 5;
            myPaneHumidity.YAxis.MinorGrid.IsVisible = true;
            myPaneHumidity.YAxis.MinorGrid.DashOn = 1;
            myPaneHumidity.YAxis.MinorGrid.DashOff = 2;
            myPaneHumidity.XAxis.MinorGrid.IsVisible = true;
            myPaneHumidity.XAxis.MinorGrid.DashOn = 1;
            myPaneHumidity.XAxis.MinorGrid.DashOff = 2;


        }
        public void SelectFoler()
        {
            path = "C:\\Users\\HP\\Documents\\" + DateTime.Now.ToString("dd.M.yyyy_HH;mm") + ".txt";
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                }
                path = fbd.SelectedPath + "\\" + DateTime.Now.ToString("dd.M.yyyy_HH;mm") + ".txt";

                selectedPath = fbd.SelectedPath;
            }
        }
        public void ReadParametrs(int adressIndex)
        {

            ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
            List<string> hexBuffer = new List<string>() { };
            for (ushort i = 2; i < 3; i++)
            {
                ushort startAddress = readAdressList[adressIndex];
                ushort numOfPoints = i;
                ushort[] holding_register = master.ReadHoldingRegisters(slaveID, startAddress,
                numOfPoints);
                Thread.Sleep(30);
                foreach (var item in holding_register)
                {
                    string intValue = item.ToString();
                    string hexValue = item.ToString("X");
                    hexBuffer.Add(hexValue);
                }
                if (hexBuffer[0] == "0")
                {
                    hexBuffer[0] = "0000";
                }
                else if (hexBuffer[1] == "0")
                {
                    hexBuffer[1] = "0000";
                }
                string hexParameter;
                if (hexBuffer[1].Length <= 1)
                {
                    hexParameter = hexBuffer[0] + hexBuffer[1] + "000";
                }
                else if (hexBuffer[1].Length <= 2)
                {
                    hexParameter = hexBuffer[0] + hexBuffer[1] + "00";
                }
                else if (hexBuffer[1].Length <= 3)
                {
                    hexParameter = hexBuffer[0] + hexBuffer[1] + "0";
                }
                else
                {
                    hexParameter = hexBuffer[0] + hexBuffer[1];
                }

                var intConvertVar = Convert.ToInt32(hexParameter, 16);
                var byteConvertVar = BitConverter.GetBytes(intConvertVar);
                float temp = BitConverter.ToSingle(byteConvertVar, 0);
                var temperature = Math.Round(temp, 2);
                if (temperature <= 0)
                {
                    readParameters.Add("0");                   
                }
                else
                {
                    readParameters.Add(Convert.ToString(temperature));
                }

                //hexList.Add(hexParameter);
            }

        }
        private void connectButton_Click_1(object sender, EventArgs e)
        {
            connectButton.Enabled = false;
            disconnectButton.Enabled = true;
            portBox.Enabled = false;
            baudBox.Enabled = false;
            slaveIDBox.Enabled = false;
            Timer1.Enabled = true;
            try
            {
                slaveID = Convert.ToByte(slaveIDBox.SelectedItem);
                serialPort1.Close();
                serialPort1.PortName = portBox.Text;
                serialPort1.BaudRate = Convert.ToInt32(baudBox.Text);
                serialPort1.Parity = Parity.None;
                serialPort1.StopBits = StopBits.One;
                serialPort1.DataBits = 8;
                serialPort1.Handshake = Handshake.None;
                serialPort1.RtsEnable = true;
                serialPort1.ReadTimeout = 500;
                serialPort1.WriteTimeout = 500;
                serialPort1.Open();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                for (int j = 0; j <= 15; j++)
                {
                    ReadParametrs(j);
                }
                TrabattoSetText.Text = readParameters[8];
                PreDryingSetTextZone1.Text = readParameters[9];
                BasicDryingSetTextZone1.Text = readParameters[11];
                BasicDryingSetTextZone2.Text = readParameters[10];
                BasicDryingSetTextZone3.Text = readParameters[12];
                BasicDryingSetTextZone4.Text = readParameters[13];
                BasicDryingSetTextZone5.Text = readParameters[14];
                //PreDryingSetTextHumidity.Text = readParameters[15];
                //BasicDryingSetTextHumidity.Text = readParameters[15];

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }
        //Timer
        //----------------------------------------------------------------------------------
        private void Timer1_Tick(object Sender, EventArgs e)
        {
            // Set the caption to the current time.  
            label2.Text = DateTime.Now.ToString();
            if (indButton == true)
            {
                buttonStatus.BackColor = Color.LimeGreen;
                indButton = false;             
            }
            else
            {
                buttonStatus.BackColor = Color.White;
                indButton = true;
            }

            try
            {
                if (serialPort1.IsOpen)
                {
                    zg1time = Convert.ToDouble(DateTime.Now.ToOADate());                 
                  
                    for (int j = 0; j <= 7; j++)
                    {
                        ReadParametrs(j);
                    }
                    //
                    TrabattoLabel.Text = Convert.ToString(Math.Round(float.Parse(readParameters[0]), 1)) + "°С";
                    PreDryingLabelZone1.Text = Convert.ToString(Math.Round(float.Parse(readParameters[1]), 1)) + "°С";
                    BasicDryingLabelZone1.Text = Convert.ToString(Math.Round(float.Parse(readParameters[3]), 1)) + "°С";
                    BasicDryingLabelZone2.Text = Convert.ToString(Math.Round(float.Parse(readParameters[2]), 1)) + "°С";
                    BasicDryingLabelZone3.Text = Convert.ToString(Math.Round(float.Parse(readParameters[4]), 1)) + "°С";
                    BasicDryingLabelZone4.Text = Convert.ToString(Math.Round(float.Parse(readParameters[5]), 1)) + "°С";
                    BasicDryingLabelZone5.Text = Convert.ToString(Math.Round(float.Parse(readParameters[6]), 1)) + "°С";
                    TrabattoLabelHumidity.Text = "n/c";
                    PreDryingLabelHumidity.Text = "n/c";
                    BasicDryingLabelHumidity.Text = "n/c";
                    //
                    TrabattoGauge.Value = float.Parse(readParameters[0]);
                    PreDryingGaugeZone1.Value = float.Parse(readParameters[1]);
                    BasicDryingGaugeZone1.Value = float.Parse(readParameters[3]);
                    BasicDryingGaugeZone2.Value = float.Parse(readParameters[2]);
                    BasicDryingGaugeZone3.Value = float.Parse(readParameters[4]);
                    BasicDryingGaugeZone4.Value = float.Parse(readParameters[5]);
                    BasicDryingGaugeZone5.Value = float.Parse(readParameters[6]);
                    TrabattoGaugeHumidity.Value = float.Parse("0");
                    PreDryingGaugeHumidity.Value = float.Parse("0");
                    BasicDryingGaugeHumidity.Value = float.Parse("0");
                    //
                    chartTemperature.Series.Clear();
                    chartTemperature.Series.Add("Temp");
                    chartTemperature.ChartAreas[0].AxisY.Maximum = 100;
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[0]));
                    chartTemperature.Series["Temp"].Points[0].Color = Color.SaddleBrown;
                    chartTemperature.Series["Temp"].Points[0].AxisLabel = "Трабатто";
                    chartTemperature.Series["Temp"].Points[0].LegendText = "Трабатто";
                    chartTemperature.Series["Temp"].Points[0].Label = (readParameters[0]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[1]));
                    chartTemperature.Series["Temp"].Points[1].Color = Color.Olive;
                    chartTemperature.Series["Temp"].Points[1].AxisLabel = "П.С.";
                    chartTemperature.Series["Temp"].Points[1].LegendText = "П.С.";
                    chartTemperature.Series["Temp"].Points[1].Label = (readParameters[1]);
                    //
                    //chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[1]));
                    //chartTemperature.Series["Temp"].Points[1].Color = Color.DarkGoldenrod;
                    //chartTemperature.Series["Temp"].Points[1].AxisLabel = "П.С. Зона 2";
                    //chartTemperature.Series["Temp"].Points[1].LegendText = "П.С.Зона 2";
                    //chartTemperature.Series["Temp"].Points[1].Label = (readParameters[1]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[3]));
                    chartTemperature.Series["Temp"].Points[2].Color = Color.Red;
                    chartTemperature.Series["Temp"].Points[2].AxisLabel = "О.С. Зона 1";
                    chartTemperature.Series["Temp"].Points[2].LegendText = "О.С. Зона 1";
                    chartTemperature.Series["Temp"].Points[2].Label = (readParameters[3]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[2]));
                    chartTemperature.Series["Temp"].Points[3].Color = Color.Blue;
                    chartTemperature.Series["Temp"].Points[3].AxisLabel = "О.С. Зона 2";
                    chartTemperature.Series["Temp"].Points[3].LegendText = "О.С. Зона 2";
                    chartTemperature.Series["Temp"].Points[3].Label = (readParameters[2]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[4]));
                    chartTemperature.Series["Temp"].Points[4].Color = Color.Green;
                    chartTemperature.Series["Temp"].Points[4].AxisLabel = "О.С. Зона 3";
                    chartTemperature.Series["Temp"].Points[4].LegendText = "О.С. Зона 3";
                    chartTemperature.Series["Temp"].Points[4].Label = (readParameters[4]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[5]));
                    chartTemperature.Series["Temp"].Points[5].Color = Color.Purple;
                    chartTemperature.Series["Temp"].Points[5].AxisLabel = "О.С. Зона 4";
                    chartTemperature.Series["Temp"].Points[5].LegendText = "О.С. Зона 4";
                    chartTemperature.Series["Temp"].Points[5].Label = (readParameters[5]);
                    //
                    chartTemperature.Series["Temp"].Points.Add(float.Parse(readParameters[6]));
                    chartTemperature.Series["Temp"].Points[6].Color = Color.Teal;
                    chartTemperature.Series["Temp"].Points[6].AxisLabel = "О.С. Зона 5";
                    chartTemperature.Series["Temp"].Points[6].LegendText = "О.С. Зона 5";
                    chartTemperature.Series["Temp"].Points[6].Label = (readParameters[6]);

                    ListPointsTrabatto.Add(new PointPair(zg1time, float.Parse(readParameters[0])));
                    ListPointsPreDrying
                        .Add(new PointPair(zg1time, float.Parse(readParameters[1])));
                    //ListPointsPreDryingZone2.Add(new PointPair(zg1time, float.Parse(readParameters[2])));
                    ListPointsBasicDryingZone1.Add(new PointPair(zg1time, float.Parse(readParameters[3])));
                    ListPointsBasicDryingZone2.Add(new PointPair(zg1time, float.Parse(readParameters[2])));
                    ListPointsBasicDryingZone3.Add(new PointPair(zg1time, float.Parse(readParameters[4])));
                    ListPointsBasicDryingZone4.Add(new PointPair(zg1time, float.Parse(readParameters[5])));
                    ListPointsBasicDryingZone5.Add(new PointPair(zg1time, float.Parse(readParameters[6])));
                    //ListPointsTrabattoHuminity.Add(new PointPair(zg1time, float.Parse(readParameters[7])));
                    //ListPointsPreDryingHuminity.Add(new PointPair(zg1time, float.Parse(readParameters[7])));
                    //ListPointsBasicDryingHuminity.Add(new PointPair(zg1time, float.Parse(readParameters[7])));
                    //
                    zedGraphControl1.AxisChange();
                    zedGraphControl1.Refresh();
                    zedGraphControl2.AxisChange();
                    zedGraphControl2.Refresh();

                    //Запис даних при запуску
                    if (saveFileInit == true)
                    {
                        number = 1;
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("№ ТР ПС ОС:1 ОС:2 ОС:3 ОС:4 ОС:5 ТР:B ПС:В ОС:В Дата Час");
                            sw.WriteLine(number.ToString() + " " + Convert.ToString(Math.Round(float.Parse(readParameters[0]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[1]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[3]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[2]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[4]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[5]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[6]), 0)) + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   DateTime.Now.ToString());
                        }
                        saveFileInit = false;
                    }
                    //Створення нового файлу при змінні поточного дня                    
                    else if (DateTime.Now.ToString("HH:mm:ss") == "08:00:00")
                    {    
                        number = 1;
                        path = selectedPath + "\\" + DateTime.Now.ToString("dd.M.yyyy_HH;mm") + ".txt";
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine("№ ТР ПС ОС:1 ОС:2 ОС:3 ОС:4 ОС:5 ТР:B ПС:В ОС:В Дата Час");
                            sw.WriteLine(number.ToString() + " " + Convert.ToString(Math.Round(float.Parse(readParameters[0]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[1]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[3]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[2]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[4]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[5]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[6]), 0)) + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   DateTime.Now.ToString());
                        }
                    }
                    //Запис даних

                    else if (DateTime.Now.ToString("mm:ss") == "00:00")
                    {
                        number += 1;
                        using (StreamWriter sw = File.AppendText(path))
                        {
                            sw.WriteLine(number.ToString() + " " + Convert.ToString(Math.Round(float.Parse(readParameters[0]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[1]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[3]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[2]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[4]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[5]), 0)) + " " +
                                                                   Convert.ToString(Math.Round(float.Parse(readParameters[6]), 0)) + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   "n/c" + " " +
                                                                   DateTime.Now.ToString());
                        }

                    }
                    readParameters.Clear();
                    hexList.Clear();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //----------------------------------------------------------------------------------

        //Set parameter (Write Register 0x10)
        //-----------------------------------------------------------------------------------
        static ushort[] SetParameter(string value)
        {
            string ToHexString(float f)
            {
                var bytes = BitConverter.GetBytes(f);
                var k = BitConverter.ToInt32(bytes, 0);
                return k.ToString("X8");
            }
            //float FromHexString(string s)
            //{
            //   var k = Convert.ToInt32(s, 16);
            //   var bytes = BitConverter.GetBytes(k);
            //  return BitConverter.ToSingle(bytes, 0);
            //}

            string hexe = ToHexString(float.Parse(value));

            string[] numberArray = new string[hexe.Length];
            int counter = 0;
            for (int v = 0; v < hexe.Length; v++)
            {
                numberArray[v] = hexe.Substring(counter, 1); // 1 is split length
                counter++;
            }
            string a = "0x40" + Convert.ToString(numberArray[0]) + Convert.ToString(numberArray[1]);
            string b = "0x" + Convert.ToString(numberArray[2]) + Convert.ToString(numberArray[3]) +
            Convert.ToString(numberArray[4]) + Convert.ToString(numberArray[5]);
            ushort h = Convert.ToUInt16(a, 16);
            ushort g = Convert.ToUInt16(b, 16);
            ushort[] datta = { h, g };
            return datta;

        }
        //-----------------------------------------------------------------------------------


        private void disconnectButton_Click(object sender, EventArgs e)
        {
            connectButton.Enabled = true;
            disconnectButton.Enabled = false;
            portBox.Enabled = true;
            baudBox.Enabled = true;
            slaveIDBox.Enabled = true;
            Timer1.Enabled = false;
            try
            {
                serialPort1.Close();
                //Get commPorts
                portBox.Items.Clear();
                string[] ports = (SerialPort.GetPortNames());
                portBox.Items.AddRange(ports);
                if (portBox.Items.Count > 0)
                {
                    portBox.SelectedIndex = 2;
                }
                 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
            }
        }

        private void basicDryingClearGraphButton_Click(object sender, EventArgs e)
        {

            ListPointsTrabatto.Clear();
            ListPointsPreDrying.Clear();
            ListPointsBasicDryingZone1.Clear();
            ListPointsBasicDryingZone2.Clear();
            ListPointsBasicDryingZone3.Clear();
            ListPointsBasicDryingZone4.Clear();
            ListPointsBasicDryingZone5.Clear();
            //ListPointsTrabattoHuminity.Clear();
            //ListPointsPreDryingHuminity.Clear();
            //ListPointsBasicDryingHuminity.Clear();
            zg1time = 0;
            zedGraphControl1.Refresh();
            zedGraphControl2.Refresh();
            graph();
        }


        private void TrabattoSetButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    //Задавати Уставку 
                    //---------------------------------------------------------------//                  
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                    ushort Address = readAdressList[8];
                    ushort[] setData = SetParameter(TrabattoSetText.Text);
                    master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------//

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void PreDryingSetButtonZone1_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    //Задавати Уставку 
                    //---------------------------------------------------------------//
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                    ushort Address = readAdressList[9];
                    ushort[] setData = SetParameter(PreDryingSetTextZone1.Text);
                    master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------//
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void BasicDryingSetButtonZone1_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    //Задавати Уставку 
                    //---------------------------------------------------------------//
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                    ushort Address = readAdressList[11];
                    ushort[] setData = SetParameter(BasicDryingSetTextZone1.Text);
                    master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------//
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BasicDryingSetButtonZone2_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                //Задавати Уставку 
                //---------------------------------------------------------------//
                ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                ushort Address = readAdressList[10];
                ushort[] setData = SetParameter(BasicDryingSetTextZone2.Text);
                master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------/
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

 
        private void BasicDryingSetButtonZone3_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                //Задавати Уставку 
                //---------------------------------------------------------------//
                ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                ushort Address = readAdressList[12];
                ushort[] setData = SetParameter(BasicDryingSetTextZone3.Text);
                master.WriteMultipleRegisters(slaveID, Address, setData);
                //----------------------------------------------------------------------------------------------//
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void BasicDryingSetButtonZone4_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                //Задавати Уставку 
                //---------------------------------------------------------------//
                ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                ushort Address = readAdressList[13];
                ushort[] setData = SetParameter(BasicDryingSetTextZone4.Text);
                master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------//
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BasicDryingSetButtonZone5_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    //Задавати Уставку 
                    //---------------------------------------------------------------//
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(serialPort1);
                    ushort Address = readAdressList[14];
                    ushort[] setData = SetParameter(BasicDryingSetTextZone5.Text);
                    master.WriteMultipleRegisters(slaveID, Address, setData);
                    //----------------------------------------------------------------------------------------------//
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void TrabattoSetButtonHumidity_Click(object sender, EventArgs e)
        {

        }
        private void PreDryingSetButtonHumidity_Click(object sender, EventArgs e)
        {

        }
        private void BasicDryingSetButtonHumidity_Click_1(object sender, EventArgs e)
        {

        }

        private void buttonReadSet_Click(object sender, EventArgs e)
        {
            try
            {
                for (int j = 0; j <= 15; j++)
                {
                    ReadParametrs(j);
                }

                TrabattoSetText.Text = readParameters[8];
                PreDryingSetTextZone1.Text = readParameters[9];
                BasicDryingSetTextZone1.Text = readParameters[11];
                BasicDryingSetTextZone2.Text = "n/c";
                BasicDryingSetTextZone3.Text = readParameters[12];
                BasicDryingSetTextZone4.Text = readParameters[13];
                BasicDryingSetTextZone5.Text = readParameters[14];
                TrabattoSetTextHumidity.Text = "n/c";
                PreDryingSetTextHumidity.Text = "n/c";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mesage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string title = "Інформація";
            string message = " Terlych Thermal Controller 2\n" +
                             " Версія програми: 1.0\n" +
                             " Назва лінії: №1\n " +
                             "Прилад: ТРМ148 v5_08(Modbus)\n " +
                             "Контакти: xeonics.technology@gmail.com\n" +
                             " © Xeonics Technology 2020.\n";
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        }


    }
}
