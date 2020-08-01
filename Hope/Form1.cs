using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using ZedGraph;
using System.IO;
using System.Runtime.CompilerServices;
using System.Data.SqlTypes;

namespace Hope
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public Microsoft.Office.Interop.Excel.Application xla;
        public Microsoft.Office.Interop.Excel.Workbook Wb;
        public Microsoft.Office.Interop.Excel.Worksheet wS;
        Microsoft.Office.Interop.Excel.Range oRng;
        
        public int i;
        public int j;
        string path;
        bool ckConnect = false;
        byte[] txbuff = new byte[64];
        byte[] RS232Buffer = new byte[64];
        double SetPoint = 0,SetPointPosition = 0 ,realSpeed = 0,Udk ,ComSetP = 0 , ComPosition, RealPosition;
        Int16 AmpMax, AmpMin;
     
        int duty;
        int dem=0;
        byte bCTHT1,bCTHT2,Green,Red;
        int state;
        int Reset=0;
        int tickStart;
        byte mode, show;
        byte[] bSetPoint = { 0, 0, 0, 0, 0, 0, 0, 0 };
        
        private void timer3_Tick(object sender, EventArgs e)
        {
            
            if (serialPort1.IsOpen)
            {

                if (serialPort1.BytesToRead > 20)
                {
                  
                    string data = serialPort1.ReadLine();
                    serialPort1.DiscardInBuffer();
                    string[] subdata = data.Split(',');

                    if (subdata.Length < 10) return;

                    if (subdata[0] == "A")
                    {

                        ComSetP = double.Parse(subdata[1]);
                        realSpeed = double.Parse(subdata[2]);
                        Udk = 0.00244*double.Parse(subdata[3]);
                        state = int.Parse(subdata[4]);
                        bCTHT1 = byte.Parse(subdata[5]);
                        bCTHT2 = byte.Parse(subdata[6]);
                        Green = byte.Parse(subdata[7]);
                        Red = byte.Parse(subdata[8]);

                        SetPointPosition = double.Parse(subdata[9]);
                        RealPosition = double.Parse(subdata[10]);

                        //if (Reset == 1 && bCTHT1 == 1)
                        //{
                        //    txbuff[59] = 0;
                        //    Reset = 0;
                        //    btStop_Click(sender, e);
                        //}

                        switch (state)
                        {
                            case 1:
                                {

                                    txtRun.Text = "Init";
                                    txtRun.ForeColor = Color.Blue;
                                    break;
                                }
                            case 2:
                                {
                                    txtRun.Text = "Reset";
                                    txtRun.ForeColor = Color.Blue;
                                    break;
                                }
                            case 3:
                                {
                                    txbuff[59] = 0;
                                    txtRun.Text = "Idle";
                                    txtRun.ForeColor = Color.Blue;
                                    break;
                                }
                            case 4:
                                {
                                    txtRun.Text = "Run";
                                    txtRun.ForeColor = Color.Green;
                                    break;
                                }
                            case 41:
                                {
                                    txtRun.Text = "Speed";
                                    txtRun.ForeColor = Color.Green;
                                    break;
                                }
                            case 42:
                                {
                                    txtRun.Text = "Position";
                                    txtRun.ForeColor = Color.Green;
                                    break;
                                }

                            case 5:
                                {
                                    txtRun.Text = "Limit";
                                    txtRun.ForeColor = Color.Blue;
                                    break;
                                }
                            case 6:
                                {
                                    txtRun.Text = "Error";
                                    txtRun.ForeColor = Color.Red;
                                    break;
                                }
                            
                        }
                         
                        if (bCTHT1 == 1)
                        {
                            txtCTHT1.Text = "ON";
                            txtCTHT1.BackColor = Color.Green;
                        }
                        else
                        {
                            txtCTHT1.Text = "OFF";
                            txtCTHT1.BackColor = Color.Red;
                        }

                        if (bCTHT2 == 1)
                        {
                            txtCTHT2.Text = "ON";
                            txtCTHT2.BackColor = Color.Green;
                        }
                        else
                        {
                            txtCTHT2.Text = "OFF";
                            txtCTHT2.BackColor = Color.Red;
                        }

                    }
                }

            }
            

        }
        private void timer1_Tick(object sender, EventArgs e)
        {        
           
        }

        
        private void timer2_Tick(object sender, EventArgs e)
        {
           // timer3.Enabled = true;
            timer5.Enabled = true;   // tính toan va gui data  
            timer2.Enabled = false;  
            timer4.Enabled = true;  // ve do thi
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
           
             ComPort.DataSource = SerialPort.GetPortNames();
            Baud.Text = "115200";
            
            string[] typesignal = { "Const", "Square", "Trapezoid" };
            SignalBox1.DataSource = typesignal;
            SignalBox2.DataSource = typesignal;
            
            //Baudrate.Text = Properties.Settings.Default.ComPortBaudrate;
           
            //KP1.Text = Properties.Settings.Default.KP1;
            //KI1.Text = Properties.Settings.Default.KI1;
            //KD1.Text = Properties.Settings.Default.KD1;

            //KP2.Text = Properties.Settings.Default.KP2;
            //KI2.Text = Properties.Settings.Default.KI2;
            //KD2.Text = Properties.Settings.Default.KD2;

            //SignalBox1.Text = Properties.Settings.Default.SignalTypePosition;
            //SignalBox2.Text = Properties.Settings.Default.SignalTypeSpeed;

            //AmMaxBox1.Text = Properties.Settings.Default.AmMax1;
            //AmMaxBox2.Text = Properties.Settings.Default.AmMax2;

            //AmMinBox1.Text = Properties.Settings.Default.AmMin1;
            //AmMinBox2.Text = Properties.Settings.Default.AmMin2;

            //show = Properties.Settings.Default.Show;

            //if (show==1)
            //{
            //    SpeedShow.Checked = true;
            //    SpeedShow.Enabled = false;
            //}
            //else if (show == 2)
            //{
            //    PositionMode.Checked = true;
            //    PositionMode.Enabled = false;
            //}
            //else
            //{
            //    UdkShow.Checked = true;
            //    UdkShow.Enabled = false;

            //}

            SpeedMode.Checked = true;
            SpeedMode.Enabled = false;
            mode = 1;

           
        }

        private void Connect_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                {
                    MessageBox.Show("Connecting is existing.");
                }
                else
                {
                    serialPort1.PortName = ComPort.Text;
                    serialPort1.BaudRate = Convert.ToInt32(Baud.Text);
                    serialPort1.Open();
                    //MessageBox.Show("Connection is successful.");
                    ckConnect = true;
                    StatusSerial.Text = "Connected";
                    StatusSerial.ForeColor = Color.Green;
                    timer5.Enabled = false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Connection is failed. " + ex.Message);
                Properties.Settings.Default.ComPortName = string.Empty;
                Form1_Load(sender, e);
            }
            if (ckConnect)
            {                              
                btStart.Enabled = true;
                btStop.Enabled = true;
            }
        }

        

        private void DisConnect_Click(object sender, EventArgs e)
        {
            try
            {
                if (txbuff[0] == 1) btStop_Click(sender, e);
                if(serialPort1.IsOpen)
                {
                    serialPort1.Close();
                    ckConnect = false;
                   // MessageBox.Show("Disconnect is successful !");                              
                    btStart.Enabled = false;btStop.Enabled = false;
                    StatusSerial.Text = "Disconnected";
                    StatusSerial.ForeColor = Color.Red;
                    timer5.Enabled = true;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Disconnect is failed. " + ex.Message); ;
            }
        }

        private void btStop_Click(object sender, EventArgs e)
        {
            txbuff[0] = 0;
            if (serialPort1.IsOpen) serialPort1.Write(txbuff, 0, 60);
            timer5.Stop();
            timer2.Stop();
            timer4.Stop();
        }

        

        private void ComPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            Properties.Settings.Default.ComPortName = ComPort.Text;
            Properties.Settings.Default.Save();
        }

        

        private void Baudrate_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.ComPortBaudrate = Baudrate.Text;
            //Properties.Settings.Default.Save();
        }
       
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {

        }

        private void SignalBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            Properties.Settings.Default.SignalTypePosition = SignalBox1.Text;
            Properties.Settings.Default.Save();
            if (SignalBox1.Text == "Const")
            {
                AmMinBox1.Enabled = false;
                PeroidBox1.Enabled = false;
            }
            else
            {
                AmMinBox1.Enabled = true;
                PeroidBox1.Enabled = true;
            }
        }

        private void AmMaxBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(AmMaxBox1.Text, out num))
                {
                    if (num > 50)
                    {
                        MessageBox.Show("Gia tri lon nhat la 50.");
                        AmMaxBox1.Text = "50";
                    }
                        
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void AmMinBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(AmMinBox1.Text, out num))
                {
                    if (num < 0)
                    {
                        MessageBox.Show("Gia tri nho nhat la 0.");
                        AmMinBox1.Text = "0";
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void PeroidBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(PeroidBox1.Text, out num))
                {
                    if(num < 0)
                    {
                        MessageBox.Show("Peroid Position phai lon hon 0");
                        PeroidBox1.Text = "5";
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void OffsetBox1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(OffsetBox1.Text, out num))
                {
                    if (num < 0)
                      MessageBox.Show("Offset phai lon hon 0");
                    OffsetBox1.Text = "0";

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
           
        }

        private void Ts1_TextChanged_1(object sender, EventArgs e)
        {
            Properties.Settings.Default.Ts1 = Ts1.Text;
            Properties.Settings.Default.Save();
        }

        private void KP1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KP1.Text, out num))
                {
                    Properties.Settings.Default.KP1 = KP1.Text;
                    Properties.Settings.Default.Save();
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void KI1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KI1.Text, out num))
                {
                    Properties.Settings.Default.KI1 = KI1.Text;
                    Properties.Settings.Default.Save();
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }

        }

        private void KD1_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KD1.Text, out num))
                {
                    Properties.Settings.Default.KD1 = KD1.Text;
                    Properties.Settings.Default.Save();
                   
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void SignalBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            Properties.Settings.Default.SignalTypeSpeed = SignalBox2.Text;
            Properties.Settings.Default.Save();
            if (SignalBox2.Text == "Const")
            {
                AmMinBox2.Enabled = false;
                PeroidBox2.Enabled = false;
            }
            else
            {
                AmMinBox2.Enabled = true;
                PeroidBox2.Enabled = true;
            }
        }

        private void AmMaxBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if(double.TryParse(AmMaxBox2.Text ,out num))
                {
                    
                    Properties.Settings.Default.AmMax2 = AmMaxBox2.Text;
                    Properties.Settings.Default.Save();

                }
                
            }
            catch(Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " +ex.Message);
            }
        }

        private void AmMinBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(AmMinBox2.Text, out num))
                {
                    
                    Properties.Settings.Default.AmMin2 = AmMinBox2.Text;
                    Properties.Settings.Default.Save();

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
        }

        private void SpeedMode_CheckedChanged_1(object sender, EventArgs e)
        {
            if (SpeedMode.Checked == true)
            {
                pictureGhep.Visible = false;
                pictureBox5.Visible = true;
                pictureBox6.Visible = true;
                mode = 1;
                SpeedMode.Enabled = false;
                PositionMode.Checked = false;
                PositionMode.Enabled = true;
                //MessageBox.Show("Mode S");

            }
        }

        private void PositionMode_CheckedChanged_1(object sender, EventArgs e)
        {
            if (PositionMode.Checked == true)
            {
                pictureGhep.Visible = true;
                pictureBox5.Visible = false;
                pictureBox6.Visible = false;
                mode = 0;
                PositionMode.Enabled = false;
                SpeedMode.Checked = false;
                SpeedMode.Enabled = true;
                // MessageBox.Show("mode P");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
             timer5.Enabled = false;
           
            
                byte[] bReset = { 1 };
                Array.Copy(bReset, 0, txbuff, 59, 1);
            
            if (serialPort1.IsOpen) serialPort1.Write(txbuff, 0, 60);
        }

        private void timer5_Tick(object sender, EventArgs e)
        {       
            if ( StatusSerial.Text != "Connected")
            {
                ComPort.DataSource = SerialPort.GetPortNames();
                timer5.Enabled = true;
            }
        }

        private void SpeedShow_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Show = 1;
            Properties.Settings.Default.Save();
            if(SpeedShow.Checked == true)
            {
                SpeedShow.Enabled = false;
                PositionShow.Checked = false;
                PositionShow.Enabled = true;
                NumSetPosition.Text = string.Empty;
                numRealPosition.Text = string.Empty;
                UdkShow.Checked = false;
                UdkShow.Enabled = true;
            }
            
            

            /******Reset Zedgraph ******/
            zedGraphControl1.GraphPane.CurveList.Clear();
            zedGraphControl1.GraphPane.GraphObjList.Clear();
            zedGraphControl1.Invalidate();
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Speed Response";
            myPane.XAxis.Title.Text = "Time (sec)";
            myPane.YAxis.Title.Text = "CM/s";
            RollingPointPairList list = new RollingPointPairList(60000);
            RollingPointPairList list1 = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("SetPoint", list, Color.Red, SymbolType.None);
            LineItem curve1 = myPane.AddCurve("RealSpeed", list1, Color.Blue, SymbolType.None);
            myPane.XAxis.Scale.Min = 0;
            myPane.XAxis.Scale.Max = 20;
            myPane.YAxis.Scale.Min = -50 + double.Parse(AmMinBox2.Text);
            myPane.YAxis.Scale.Max = +50 + double.Parse(AmMaxBox2.Text);
            myPane.YAxis.Scale.MinorStep = 1;
            myPane.YAxis.Scale.MajorStep = 2;
            myPane.XAxis.Scale.MinorStep = 1;
            myPane.XAxis.Scale.MajorStep = 5;
            zedGraphControl1.AxisChange();
            tickStart = Environment.TickCount;
        }

        private void SetValue_Click(object sender, EventArgs e)
        {
            
            if (SpeedMode.Checked)
            {

                double num;

                if (!Int16.TryParse(AmMaxBox2.Text, out AmpMax))
                {
                    MessageBox.Show("Am Max Speed  phai la 1 so !");
                    return;
                }
                if (!Int16.TryParse(AmMinBox2.Text, out AmpMin))
                {
                    MessageBox.Show("Am Min Speed  phai la 1 so !");
                    return;
                }
                if (SignalBox2.Text == "Square" || SignalBox2.Text == "Trapezoid")
                {
                    if (AmpMax > 100)
                    {
                        //MessageBox.Show("AmpMax <= 100 !");
                        AmMaxBox2.Text = "100";
                        return;
                    }
                    if (AmpMin < -100)
                    {
                        //MessageBox.Show("Am Min >= -100 !");
                        AmMinBox2.Text = "-100";
                        return;
                    }
                }
                

                if (!double.TryParse(PeroidBox2.Text, out num))
                {
                    MessageBox.Show("Peroid Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(OffsetBox2.Text, out num))
                {
                    MessageBox.Show("Offset Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KP2.Text, out num))
                {
                    MessageBox.Show("KP PId Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KI2.Text, out num))
                {
                    MessageBox.Show("KI PId Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KD2.Text, out num))
                {
                    MessageBox.Show("KD PId Speed phai la 1 so !");
                    return;
                }
                txbuff[1] = 1;
                byte[] bAmmax = BitConverter.GetBytes(AmpMax);
                Array.Copy(bAmmax, 0, txbuff, 52, 2);
                byte[] bAmmin = BitConverter.GetBytes(AmpMin);
                Array.Copy(bAmmin, 0, txbuff, 54, 2);
                byte[] Period2 = BitConverter.GetBytes(Int16.Parse(PeroidBox2.Text));
                Array.Copy(Period2, 0, txbuff, 56, 2);

                if (SignalBox2.Text == "Const")
                {
                    byte[] bmodspeed = { 1 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox2.Text == "Square")
                {
                    byte[] bmodspeed = { 2 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox2.Text == "Trapezoid")
                {
                    byte[] bmodspeed = { 3 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }

                byte[] bKp2 = BitConverter.GetBytes(double.Parse(KP2.Text));
                Array.Copy(bKp2, 0, txbuff, 28, 8);
                byte[] bKi2 = BitConverter.GetBytes(double.Parse(KI2.Text));
                Array.Copy(bKi2, 0, txbuff, 36, 8);
                byte[] bKd2 = BitConverter.GetBytes(double.Parse(KD2.Text));
                Array.Copy(bKd2, 0, txbuff, 44, 8);
            }

            if (PositionMode.Checked)
            {

                double num;
                if (!Int16.TryParse(AmMaxBox1.Text, out AmpMax))
                {
                    MessageBox.Show("Am Max Position  phai la 1 so !");
                    return;
                }
                if (!Int16.TryParse(AmMinBox1.Text, out AmpMin))
                {
                    MessageBox.Show("Am Min Position  phai la 1 so !");
                    return;
                }
{
                    if (AmpMax > 50)
                    {
                       // MessageBox.Show("AmpMax <= 50 !");
                        AmMaxBox1.Text = "50";
                        return;
                    }                 
                }

                if (!double.TryParse(PeroidBox1.Text, out num))
                {
                    MessageBox.Show("Peroid Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(OffsetBox1.Text, out num))
                {
                    MessageBox.Show("Offset Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KP1.Text, out num))
                {
                    MessageBox.Show("KP PId Position phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KI1.Text, out num))
                {
                    MessageBox.Show("KI PId Position phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KD1.Text, out num))
                {
                    MessageBox.Show("KD PId Position phai la 1 so !");
                    return;
                }
                txbuff[1] = 0;
                byte[] bAmmax = BitConverter.GetBytes(AmpMax);
                Array.Copy(bAmmax, 0, txbuff, 52, 2);
                byte[] bAmmin = BitConverter.GetBytes(AmpMin);
                Array.Copy(bAmmin, 0, txbuff, 54, 2);
                byte[] Period1 = BitConverter.GetBytes(Int16.Parse(PeroidBox1.Text));
                Array.Copy(Period1, 0, txbuff, 56, 2);

                if (SignalBox1.Text == "Const")
                {
                    byte[] bmodspeed = { 1 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox1.Text == "Square")
                {
                    byte[] bmodspeed = { 2 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox1.Text == "Trapezoid")
                {
                    byte[] bmodspeed = { 3 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }

                byte[] bKp1 = BitConverter.GetBytes(double.Parse(KP1.Text));
                Array.Copy(bKp1, 0, txbuff, 3, 8);
                byte[] bKi1 = BitConverter.GetBytes(double.Parse(KI1.Text));
                Array.Copy(bKi1, 0, txbuff, 11, 8);
                byte[] bKD1 = BitConverter.GetBytes(double.Parse(KD1.Text));
                Array.Copy(bKD1, 0, txbuff, 11, 8);
            }
            MessageBox.Show("Done Set Value!");

            if (serialPort1.IsOpen) serialPort1.Write(txbuff, 0, 60);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SetValue_Click(sender, e);
                MessageBox.Show("Set Value");
            }
        }

        private void txtRun_Click(object sender, EventArgs e)
        {

        }

        private void PositionShow_CheckedChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Show = 2;
            //Properties.Settings.Default.Save();

            if (PositionShow.Checked)
            {
                SpeedShow.Checked = false;
                SpeedShow.Enabled = true;
                UdkShow.Checked = false;
                UdkShow.Enabled = true;
                PositionShow.Enabled = false;
            }
            

            /******Reset Zedgraph ******/
            zedGraphControl1.GraphPane.CurveList.Clear();
            zedGraphControl1.GraphPane.GraphObjList.Clear();
            zedGraphControl1.Invalidate();
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Position Response";
            myPane.XAxis.Title.Text = "Time (sec)";
            myPane.YAxis.Title.Text = "Cm";
            RollingPointPairList list = new RollingPointPairList(60000);
            RollingPointPairList list1 = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("SetPoint", list, Color.Red, SymbolType.None);
            LineItem curve1 = myPane.AddCurve("Real Position", list1, Color.Blue, SymbolType.None);
            myPane.XAxis.Scale.Min = 0;
            myPane.XAxis.Scale.Max = 20;
            myPane.YAxis.Scale.Min = -50 + double.Parse(AmMinBox2.Text);
            myPane.YAxis.Scale.Max = +50 + double.Parse(AmMaxBox2.Text);
            myPane.YAxis.Scale.MinorStep = 1;
            myPane.YAxis.Scale.MajorStep = 2;
            myPane.XAxis.Scale.MinorStep = 1;
            myPane.XAxis.Scale.MajorStep = 5;
            zedGraphControl1.AxisChange();
            tickStart = Environment.TickCount;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            path = DateTime.Now.ToString("hh.mm.ss_dd.MM.yyy") + ".xlsx";
            xla = new Microsoft.Office.Interop.Excel.Application();
            Wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            wS = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;

            Microsoft.Office.Interop.Excel.Range rg = (Microsoft.Office.Interop.Excel.Range)wS.get_Range("A1", "F1");
            wS.Cells[1, 1] = "Times";
            wS.Cells[1, 2] = "Set Speed";
            wS.Cells[1, 3] = "Real Speed";
            wS.Cells[1, 4] = "Set Position";
            wS.Cells[1, 5] = "Real Position";
            wS.Cells[1, 6] = "Udk";
            rg.Columns.AutoFit();

            i = 2; j = 1;
            

        }

        private void UdkShow_CheckedChanged(object sender, EventArgs e)
        {
            //Properties.Settings.Default.Show = 3;
            //Properties.Settings.Default.Save();

            if (UdkShow.Enabled)
            {
                UdkShow.Enabled = false;
                SpeedShow.Checked = false;
                SpeedShow.Enabled = true;
                PositionShow.Checked = false;
                PositionShow.Enabled = true;

            }


            zedGraphControl1.GraphPane.CurveList.Clear();
            zedGraphControl1.GraphPane.GraphObjList.Clear();
            zedGraphControl1.Invalidate();
            GraphPane myPane = zedGraphControl1.GraphPane;
            myPane.Title.Text = "Momen Set";
            myPane.XAxis.Title.Text = "Time (sec)";
            myPane.YAxis.Title.Text = "Voltage";
            RollingPointPairList list = new RollingPointPairList(60000);
            LineItem curve = myPane.AddCurve("Udk", list, Color.Red, SymbolType.None);
           
            myPane.XAxis.Scale.Min = 0;
            myPane.XAxis.Scale.Max = 20;
            myPane.YAxis.Scale.Min = -50 + double.Parse(AmMinBox2.Text);
            myPane.YAxis.Scale.Max = +50 + double.Parse(AmMaxBox2.Text);
            myPane.YAxis.Scale.MinorStep = 1;
            myPane.YAxis.Scale.MajorStep = 2;
            myPane.XAxis.Scale.MinorStep = 1;
            myPane.XAxis.Scale.MajorStep = 5;
            zedGraphControl1.AxisChange();
            tickStart = Environment.TickCount;
        }

        private void PeroidBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(PeroidBox2.Text, out num))
                {
                    if(num < 0)
                    {
                        MessageBox.Show("Peroid Speed phai lon hon 0");
                        PeroidBox2.Text = "0";
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void OffsetBox2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(OffsetBox2.Text, out num))
                {
                    if (num < 0)
                    {
                        MessageBox.Show("Offset phai lon hon 0");
                        OffsetBox2.Text = "0";
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void KP2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KP2.Text, out num))
                {
                    Properties.Settings.Default.KP2 = KP2.Text;
                    Properties.Settings.Default.Save();
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void KI2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KI2.Text, out num))
                {
                    Properties.Settings.Default.KI2 = KI2.Text;
                    Properties.Settings.Default.Save();
                   
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }

           
        }

        private void KD2_TextChanged_1(object sender, EventArgs e)
        {
            try
            {
                double num;
                if (double.TryParse(KD2.Text, out num))
                {
                    Properties.Settings.Default.KD2 = KD2.Text;
                    Properties.Settings.Default.Save();
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui long nhap 1 so. " + ex.Message);
            }
            
        }

        private void Ts2_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.Ts2 = Ts2.Text;
            Properties.Settings.Default.Save();
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            if (SpeedShow.Checked)
            {
                Draw(ComSetP, realSpeed);
                if (SpeedShow.Checked)
                {
                    dothiSetPoint.Text = ComSetP.ToString();
                    dothiReal.Text = realSpeed.ToString();
                    NumSetPosition.Text = string.Empty;
                    numRealPosition.Text = string.Empty;
                }
                
                Udktxt.Text = Udk.ToString();
                
            }
             if (PositionShow.Checked)
            {
                Draw(SetPointPosition, RealPosition);
                if (PositionMode.Checked)
                {
                    NumSetPosition.Text = SetPointPosition.ToString();
                    numRealPosition.Text = RealPosition.ToString();
                    dothiSetPoint.Text = string.Empty;
                    dothiReal.Text = string.Empty;
                    Udktxt.Text = Udk.ToString();
                }
            }
             if (UdkShow.Checked)
            {
                Udk_Draw(Udk);
            }
           // 
           
            
        }

        private void btStart_Click(object sender, EventArgs e)
        {

            if (SpeedMode.Checked)
            {

                double num;

                if (!Int16.TryParse(AmMaxBox2.Text, out AmpMax))
                {
                    MessageBox.Show("Am Max Speed  phai la 1 so !");
                    return;
                }
                if (!Int16.TryParse(AmMinBox2.Text, out AmpMin))
                {
                    MessageBox.Show("Am Min Speed  phai la 1 so !");
                    return;
                }
                if (SignalBox2.Text == "Square" || SignalBox2.Text == "Tropezoid")
                {
                    if (AmpMax > 100)
                    {
                       // MessageBox.Show("AmpMax <= 100 !");
                        AmMaxBox2.Text = "100";
                        return;
                    }
                    if (AmpMin < -100)
                    {
                        //MessageBox.Show("Am Min >= -100 !");
                        AmMinBox2.Text = "-100";
                        return;
                    }
                }
                

                if (!double.TryParse(PeroidBox2.Text, out num))
                {
                    MessageBox.Show("Peroid Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(OffsetBox2.Text, out num))
                {
                    MessageBox.Show("Offset Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KP2.Text, out num))
                {
                    MessageBox.Show("KP PId Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KI2.Text, out num))
                {
                    MessageBox.Show("KI PId Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KD2.Text, out num))
                {
                    MessageBox.Show("KD PId Speed phai la 1 so !");
                    return;
                }
                txbuff[1] = 1;
                byte[] bAmmax = BitConverter.GetBytes(AmpMax);
                Array.Copy(bAmmax, 0, txbuff, 52, 2);
                byte[] bAmmin = BitConverter.GetBytes(AmpMin);
                Array.Copy(bAmmin, 0, txbuff, 54, 2);
                byte[] Period2 = BitConverter.GetBytes(Int16.Parse(PeroidBox2.Text));
                Array.Copy(Period2, 0, txbuff, 56, 2);

                if (SignalBox2.Text == "Const")
                {
                    byte[] bmodspeed = { 1 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox2.Text == "Square")
                {
                    byte[] bmodspeed = { 2 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox2.Text == "Trapezoid")
                {
                    byte[] bmodspeed = { 3 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                duty = int.Parse(PeroidBox2.Text);
                if (OffsetBox2.Text != "0")
                {
                    timer2.Interval = 1000 * int.Parse(OffsetBox2.Text);
                    timer2.Enabled = true;
                }
                else
                {
                    timer1.Enabled = true;
                    timer4.Enabled = true;
                }
                     

                /******Reset Zedgraph ******/
                zedGraphControl1.GraphPane.CurveList.Clear();
                zedGraphControl1.GraphPane.GraphObjList.Clear();
                zedGraphControl1.Invalidate();
                GraphPane myPane = zedGraphControl1.GraphPane;
                myPane.Title.Text = "Speed Response";
                myPane.XAxis.Title.Text = "Time (sec)";
                myPane.YAxis.Title.Text = "Cm/s";
                RollingPointPairList list = new RollingPointPairList(60000);
                RollingPointPairList list1 = new RollingPointPairList(60000);
                LineItem curve = myPane.AddCurve("SetPoint", list, Color.Red, SymbolType.None);
                LineItem curve1 = myPane.AddCurve("RealSpeed", list1, Color.Blue, SymbolType.None);
                myPane.XAxis.Scale.Min = 0;
                myPane.XAxis.Scale.Max = 20;
                myPane.YAxis.Scale.Min = -50 + double.Parse(AmMinBox2.Text);
                myPane.YAxis.Scale.Max = +50 + double.Parse(AmMaxBox2.Text);
                myPane.YAxis.Scale.MinorStep = 1;
                myPane.YAxis.Scale.MajorStep = 2;
                myPane.XAxis.Scale.MinorStep = 1;
                myPane.XAxis.Scale.MajorStep = 5;
                zedGraphControl1.AxisChange();
                tickStart = Environment.TickCount;
            } 
                
            if (PositionMode.Checked)
            {
                double num;
                if (!Int16.TryParse(AmMaxBox1.Text, out AmpMax))
                {
                    MessageBox.Show("Am Max Position  phai la 1 so !");
                    return;
                }
                if (!Int16.TryParse(AmMinBox1.Text, out AmpMin))
                {
                    MessageBox.Show("Am Min Position  phai la 1 so !");
                    return;
                }
              
                if (SignalBox1.Text == "Square" || SignalBox1.Text == "Tropezoid")
                {
                    if (AmpMax > 50)
                    {
                       // MessageBox.Show("AmpMax <= 50 !");
                        AmMaxBox1.Text = "50";
                        return;
                    }
                    
                }
                if (!double.TryParse(PeroidBox1.Text, out num))
                {
                    MessageBox.Show("Peroid Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(OffsetBox1.Text, out num))
                {
                    MessageBox.Show("Offset Speed phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KP1.Text, out num))
                {
                    MessageBox.Show("KP PId Position phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KI1.Text, out num))
                {
                    MessageBox.Show("KI PId Position phai la 1 so !");
                    return;
                }
                if (!double.TryParse(KD1.Text, out num))
                {
                    MessageBox.Show("KD PId Position phai la 1 so !");
                    return;
                }
                txbuff[1] = 0;
                byte[] bAmmax = BitConverter.GetBytes(AmpMax);
                Array.Copy(bAmmax, 0, txbuff, 52, 2);
                byte[] bAmmin = BitConverter.GetBytes(AmpMin);
                Array.Copy(bAmmin, 0, txbuff, 54, 2);
                byte[] Period1 = BitConverter.GetBytes(Int16.Parse(PeroidBox1.Text));
                Array.Copy(Period1, 0, txbuff, 56, 2);

                if (SignalBox1.Text == "Const")
                {
                    byte[] bmodspeed = { 1 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox1.Text == "Square")
                {
                    byte[] bmodspeed = { 2 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                if (SignalBox1.Text == "Trapezoid")
                {
                    byte[] bmodspeed = { 3 };
                    Array.Copy(bmodspeed, 0, txbuff, 58, 1);
                }
                duty = int.Parse(PeroidBox1.Text);
                if (OffsetBox1.Text != "0")
                {
                    timer2.Interval = 1000 * int.Parse(OffsetBox1.Text);
                    timer2.Enabled = true;
                }
                else
                {
                    timer1.Enabled = true;
                    timer4.Enabled = true;
                }


                /******Reset Zedgraph ******/
                zedGraphControl1.GraphPane.CurveList.Clear();
                zedGraphControl1.GraphPane.GraphObjList.Clear();
                zedGraphControl1.Invalidate();
                GraphPane myPane = zedGraphControl1.GraphPane;
                myPane.Title.Text = "Position Response";
                myPane.XAxis.Title.Text = "Time (sec)";
                myPane.YAxis.Title.Text = "Cm";
                RollingPointPairList list = new RollingPointPairList(60000);
                RollingPointPairList list1 = new RollingPointPairList(60000);
                LineItem curve = myPane.AddCurve("SetPoint", list, Color.Red, SymbolType.None);
                LineItem curve1 = myPane.AddCurve("RealPosition", list1, Color.Blue, SymbolType.None);
                myPane.XAxis.Scale.Min = 0;
                myPane.XAxis.Scale.Max = 20;
                myPane.YAxis.Scale.Min = -50 + double.Parse(AmMinBox1.Text);
                myPane.YAxis.Scale.Max = +50 + double.Parse(AmMaxBox1.Text);
                myPane.YAxis.Scale.MinorStep = 1;
                myPane.YAxis.Scale.MajorStep = 2;
                myPane.XAxis.Scale.MinorStep = 1;
                myPane.XAxis.Scale.MajorStep = 5;
                zedGraphControl1.AxisChange();
                tickStart = Environment.TickCount;
            }
                

                /*** Du lieu*////
                //Byte : 0   |1   |2  |3-10|11-18|19-26|27 |28-35|36-43|44-51|52-53|54-55|56  -57   |58      |59   |
                // Start/stop|Mode|Ts1|Kp1 |Ki1  |Kd1  |Ts2|Kp2  |Ki2  |Kd2  |Am-MAx|AmMin|Peroid|Signal type|Reset|

                
                txbuff[0] = 1;
                txbuff[1] = mode;
                txbuff[2] = byte.Parse(Ts1.Text);
                txbuff[27] = byte.Parse(Ts2.Text);
                byte[] bKp1 = BitConverter.GetBytes(double.Parse(KP1.Text));
                Array.Copy(bKp1, 0, txbuff, 3, 8);

                byte[] bKi1 = BitConverter.GetBytes(double.Parse(KI1.Text));
                Array.Copy(bKi1, 0, txbuff, 11, 8);

                byte[] bKd1 = BitConverter.GetBytes(double.Parse(KD1.Text));
                Array.Copy(bKd1, 0, txbuff, 19, 8);

                byte[] bKp2 = BitConverter.GetBytes(double.Parse(KP2.Text));
                Array.Copy(bKp2, 0, txbuff, 28, 8);

                byte[] bKi2 = BitConverter.GetBytes(double.Parse(KI2.Text));
                Array.Copy(bKi2, 0, txbuff, 36, 8);

                byte[] bKd2 = BitConverter.GetBytes(double.Parse(KD2.Text));
                Array.Copy(bKd2, 0, txbuff, 44, 8);

                //byte[] bSetPoint = BitConverter.GetBytes(SetPoint);
                //Array.Copy(bSetPoint, 0, txbuff, 52, 8);

            

            if (serialPort1.IsOpen) serialPort1.Write(txbuff, 0, 60);

        }

        private void Draw(double setpoint,double real)
        {        
            
          if (zedGraphControl1.GraphPane.CurveList.Count <= 0)
                return;

            LineItem curve = zedGraphControl1.GraphPane.CurveList[0] as LineItem;
            LineItem curve1 = zedGraphControl1.GraphPane.CurveList[1] as LineItem;
            if (curve == null) return;
            if (curve1 == null) return;

            IPointListEdit list = curve.Points as IPointListEdit;
            IPointListEdit list1 = curve1.Points as IPointListEdit;

            if (list == null) return;
            if (list1 == null) return;

            double time = (Environment.TickCount - tickStart) / 1000.0;

            list.Add(time,setpoint);
            list1.Add(time,real);
            Scale xscale = zedGraphControl1.GraphPane.XAxis.Scale;
            if (time > xscale.Max - xscale.MajorStep)
            {
                xscale.Max = time + xscale.MajorStep;
                xscale.Min = xscale.Max - 20.0;
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            zedGraphControl1.Refresh();
        }

        private void Udk_Draw(double value)
        {

            
            if (zedGraphControl1.GraphPane.CurveList.Count <= 0)
                return;

            LineItem curve = zedGraphControl1.GraphPane.CurveList[0] as LineItem;
            if (curve == null) return;
            IPointListEdit list = curve.Points as IPointListEdit;      
            if (list == null) return;           
            double time = (Environment.TickCount - tickStart) / 1000.0;

            list.Add(time, value);           
            Scale xscale = zedGraphControl1.GraphPane.XAxis.Scale;
            if (time > xscale.Max - xscale.MajorStep)
            {
                xscale.Max = time + xscale.MajorStep;
                xscale.Min = xscale.Max - 20.0;
            }
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
            zedGraphControl1.Refresh();
        }
        

       
    }
}
