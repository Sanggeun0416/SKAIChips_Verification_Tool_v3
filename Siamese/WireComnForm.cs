#region 어셈블리 JLcLib, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// G:\Verification_Tools\SKAIChips_TestGUI_V3.1.1_Source\Siamese\bin\Debug\JLcLib.dll
// Decompiled with ICSharpCode.Decompiler 8.2.0.7535
#endregion

using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO.Ports;
using System.Windows.Forms;
using JLcLib.Comn;

namespace JLcLib.Custom
{
    public class WireComnForm : Form
    {
        private int[] UART_BaudRates = new int[16]
        {
            110, 300, 1200, 2400, 4800, 9600, 19200, 38400, 50000, 57800,
            115200, 230400, 460800, 921600, 2000000, 2083312
        };

        private int[] UART_DataBits = new int[2] { 7, 8 };

        private Serial UART = null;

        private SPI SPI = null;

        private I2C I2C = null;

        private IContainer components = null;

        private TabControl Comn_TabControl;

        private TabPage Serial_TabPage;

        private TabPage SPI_TabPage;

        private TabPage I2C_TabPage;

        private ComboBox UART_FlowCtrl_ComboBox;

        private ComboBox UART_StopBit_ComboBox;

        private ComboBox UART_Parity_ComboBox;

        private ComboBox UART_DataBits_ComboBox;

        private ComboBox UART_BaudRate_ComboBox;

        private ComboBox UART_Port_ComboBox;

        private Label label1;

        private Label label6;

        private Label label5;

        private Label label4;

        private Label label3;

        private Label label2;

        private Button SerialOK_Button;

        private Button SerialCancel_Button;

        private Button SPI_Cancel_Button;

        private Button SPI_OK_Button;

        private Label label7;

        private ComboBox SPI_Mode_ComboBox;

        private Label label10;

        private ComboBox SPI_BitFirst_ComboBox;

        private Label label8;

        private TextBox SPI_Clock_TextBox;

        private ComboBox SPI_Device_ComboBox;

        private Label label16;

        private Label label17;

        private ComboBox I2C_Device_ComboBox;

        private TextBox I2C_SlaveAddr_TextBox;

        private Label label24;

        private TextBox I2C_Clock_TextBox;

        private Label label26;

        private Button I2C_Cancel_Button;

        private Button I2C_OK_Button;

        private Button SPI_Refresh_Button;

        private Button I2C_Refresh_Button;

        private Button UART_Refresh_Button;

        public WireComnTypes ComnType { get; set; }

        public WireComnForm(object ComnClass)
        {
            InitializeComponent();
            if (ComnClass.GetType() == typeof(Serial))
            {
                UART = (Serial)ComnClass;
            }
            else if (ComnClass.GetType() == typeof(SPI))
            {
                SPI = (SPI)ComnClass;
            }
            else if (ComnClass.GetType() == typeof(I2C))
            {
                I2C = (I2C)ComnClass;
            }
        }

        private void WireComnForm_Load(object sender, EventArgs e)
        {
            Comn_TabControl.TabPages.Clear();
            if (UART != null)
            {
                Comn_TabControl.TabPages.Add(Serial_TabPage);
                UART_SetUI();
            }

            if (SPI != null)
            {
                Comn_TabControl.TabPages.Add(SPI_TabPage);
                SPI_SetUI();
            }

            if (I2C != null)
            {
                Comn_TabControl.TabPages.Add(I2C_TabPage);
                I2C_SetUI();
            }

            Comn_TabControl.SelectedIndex = 0;
        }

        private void WireComnForm_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void Comn_TabControl_Selected(object sender, TabControlEventArgs e)
        {
        }

        public void SelectTab(WireComnTypes Type)
        {
            Comn_TabControl.SelectedIndex = (int)Type;
        }

        public static void Show(object ComnClass)
        {
            WireComnForm wireComnForm = new WireComnForm(ComnClass);
            wireComnForm.ShowDialog();
        }

        private void UART_SetUI()
        {
            foreach (string devicesName in UART.DevicesNames)
            {
                UART_Port_ComboBox.Items.Add(devicesName);
            }

            UART_Port_ComboBox.SelectedItem = UART.Config.PortName;
            if (UART_Port_ComboBox.SelectedIndex == -1 && UART_Port_ComboBox.Items.Count > 0)
            {
                UART_Port_ComboBox.SelectedIndex = 0;
            }

            int[] uART_BaudRates = UART_BaudRates;
            foreach (int num in uART_BaudRates)
            {
                UART_BaudRate_ComboBox.Items.Add(num.ToString());
            }

            UART_BaudRate_ComboBox.SelectedItem = UART.Config.BaudRate.ToString();
            int[] uART_DataBits = UART_DataBits;
            foreach (int num2 in uART_DataBits)
            {
                UART_DataBits_ComboBox.Items.Add(num2 + " Bit");
            }

            UART_DataBits_ComboBox.SelectedIndex = UART.Config.DataBit - UART_DataBits[0];
            UART_StopBit_ComboBox.Items.Add("1 Bit");
            UART_StopBit_ComboBox.Items.Add("2 Bit");
            UART_StopBit_ComboBox.Items.Add("1.5 Bit");
            UART_StopBit_ComboBox.SelectedIndex = (int)(UART.Config.StopBit - 1);
            for (int k = 0; k <= 4; k++)
            {
                ComboBox.ObjectCollection items = UART_Parity_ComboBox.Items;
                Parity parity = (Parity)k;
                items.Add(parity.ToString());
            }

            UART_Parity_ComboBox.SelectedIndex = (int)UART.Config.Parity;
            UART_FlowCtrl_ComboBox.Items.Add("None");
            UART_FlowCtrl_ComboBox.Items.Add("Xon/Xoff");
            UART_FlowCtrl_ComboBox.Items.Add("Hardware");
            UART_FlowCtrl_ComboBox.Items.Add("Xon/Xoff Hardware");
            UART_FlowCtrl_ComboBox.SelectedIndex = (int)UART.Config.FlowCont;
        }

        private void UART_StoreSetting()
        {
            UART.Config.PortName = UART_Port_ComboBox.SelectedItem.ToString();
            UART.Config.BaudRate = UART_BaudRates[UART_BaudRate_ComboBox.SelectedIndex];
            UART.Config.DataBit = UART_DataBits[UART_DataBits_ComboBox.SelectedIndex];
            UART.Config.StopBit = (StopBits)(UART_StopBit_ComboBox.SelectedIndex + 1);
            UART.Config.Parity = (Parity)UART_Parity_ComboBox.SelectedIndex;
            UART.Config.FlowCont = (Handshake)UART_FlowCtrl_ComboBox.SelectedIndex;
        }

        private void UART_Refresh_Button_Click(object sender, EventArgs e)
        {
            UART_Port_ComboBox.Items.Clear();
            foreach (string item in UART.Search())
            {
                UART_Port_ComboBox.Items.Add(item);
            }

            UART_Port_ComboBox.SelectedItem = UART.Config.PortName;
            if (UART_Port_ComboBox.SelectedIndex == -1 && UART_Port_ComboBox.Items.Count > 0)
            {
                UART_Port_ComboBox.SelectedIndex = 0;
            }
        }

        private void UART_OK_Button_Click(object sender, EventArgs e)
        {
            UART_StoreSetting();
            Close();
        }

        private void UART_Cancel_Button_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void SPI_SetUI()
        {
            foreach (string devicesName in SPI.DevicesNames)
            {
                SPI_Device_ComboBox.Items.Add(devicesName);
            }

            if (SPI_Device_ComboBox.Items.Count > 0)
            {
                if (SPI.Config.DevIndex >= 0 && SPI.Config.DevIndex < SPI_Device_ComboBox.Items.Count)
                {
                    if (SPI.Config.DevName == (string)SPI_Device_ComboBox.Items[SPI.Config.DevIndex])
                    {
                        SPI_Device_ComboBox.SelectedIndex = SPI.Config.DevIndex;
                    }
                    else
                    {
                        SPI_Device_ComboBox.SelectedItem = SPI.Config.DevName;
                    }
                }
                else
                {
                    SPI_Device_ComboBox.SelectedItem = SPI.Config.DevName;
                }

                if (SPI_Device_ComboBox.SelectedIndex == -1)
                {
                    SPI_Device_ComboBox.SelectedIndex = 0;
                }
            }

            SPI_Clock_TextBox.Text = SPI.Config.Clock_kHz.ToString();
            SPI_Mode_ComboBox.Items.Add("CPOL=0,CPHA=0");
            SPI_Mode_ComboBox.Items.Add("CPOL=0,CPHA=1");
            SPI_Mode_ComboBox.Items.Add("CPOL=1,CPHA=0");
            SPI_Mode_ComboBox.Items.Add("CPOL=1,CPHA=1");
            SPI_Mode_ComboBox.SelectedIndex = SPI.Config.Mode;
            if (SPI_Mode_ComboBox.SelectedIndex == -1 && SPI_Mode_ComboBox.Items.Count > 0)
            {
                SPI_Mode_ComboBox.SelectedIndex = 0;
            }

            SPI_BitFirst_ComboBox.Items.Add("MSB");
            SPI_BitFirst_ComboBox.Items.Add("LSB");
            SPI_BitFirst_ComboBox.SelectedIndex = SPI.Config.BitFirst;
            if (SPI_BitFirst_ComboBox.SelectedIndex == -1 && SPI_BitFirst_ComboBox.Items.Count > 0)
            {
                SPI_BitFirst_ComboBox.SelectedIndex = 0;
            }
        }

        private void SPI_StoreSetting()
        {
            SPI.Config.DevName = SPI_Device_ComboBox.SelectedItem.ToString();
            SPI.Config.DevIndex = SPI_Device_ComboBox.SelectedIndex;
            SPI.Config.Clock_kHz = (int.TryParse(SPI_Clock_TextBox.Text, out var result) ? result : SPI.Config.Clock_kHz);
            SPI.Config.Mode = SPI_Mode_ComboBox.SelectedIndex;
            SPI.Config.BitFirst = SPI_BitFirst_ComboBox.SelectedIndex;
        }

        private void SPI_Refresh_Button_Click(object sender, EventArgs e)
        {
            SPI_Device_ComboBox.Items.Clear();
            foreach (string item in SPI.Search())
            {
                SPI_Device_ComboBox.Items.Add(item);
            }

            SPI_Device_ComboBox.SelectedItem = SPI.Config.DevName;
            if (SPI_Device_ComboBox.SelectedIndex == -1 && SPI_Device_ComboBox.Items.Count > 0)
            {
                SPI_Device_ComboBox.SelectedIndex = 0;
            }
        }

        private void SPI_OK_Button_Click(object sender, EventArgs e)
        {
            SPI_StoreSetting();
            Close();
        }

        private void SPI_Cancel_Button_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void I2C_SetUI()
        {
            foreach (string devicesName in I2C.DevicesNames)
            {
                I2C_Device_ComboBox.Items.Add(devicesName);
            }

            if (I2C_Device_ComboBox.Items.Count > 0)
            {
                if (I2C.Config.DevIndex < I2C_Device_ComboBox.Items.Count && I2C.Config.DevName == (string)I2C_Device_ComboBox.Items[I2C.Config.DevIndex])
                {
                    I2C_Device_ComboBox.SelectedIndex = I2C.Config.DevIndex;
                }
                else
                {
                    I2C_Device_ComboBox.SelectedItem = I2C.Config.DevName;
                }

                if (I2C_Device_ComboBox.SelectedIndex == -1)
                {
                    I2C_Device_ComboBox.SelectedIndex = 0;
                }
            }

            I2C_Clock_TextBox.Text = I2C.Config.Clock_kHz.ToString();
            I2C_SlaveAddr_TextBox.Text = I2C.Config.SlaveAddress.ToString("X2");
        }

        private void I2C_StoreSetting()
        {
            I2C.Config.DevName = I2C_Device_ComboBox.SelectedItem.ToString();
            I2C.Config.DevIndex = I2C_Device_ComboBox.SelectedIndex;
            I2C.Config.Clock_kHz = (int.TryParse(I2C_Clock_TextBox.Text, out var result) ? result : I2C.Config.Clock_kHz);
            try
            {
                I2C.Config.SlaveAddress = int.Parse(I2C_SlaveAddr_TextBox.Text, NumberStyles.HexNumber);
            }
            catch
            {
            }
        }

        private void I2C_Refresh_Button_Click(object sender, EventArgs e)
        {
            I2C_Device_ComboBox.Items.Clear();
            foreach (string item in I2C.Search())
            {
                I2C_Device_ComboBox.Items.Add(item);
            }

            I2C_Device_ComboBox.SelectedItem = I2C.Config.DevName;
            if (I2C_Device_ComboBox.SelectedIndex == -1 && I2C_Device_ComboBox.Items.Count > 0)
            {
                I2C_Device_ComboBox.SelectedIndex = 0;
            }
        }

        private void I2C_OK_Button_Click(object sender, EventArgs e)
        {
            I2C_StoreSetting();
            Close();
        }

        private void I2C_Cancel_Button_Click(object sender, EventArgs e)
        {
            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.Comn_TabControl = new System.Windows.Forms.TabControl();
            this.Serial_TabPage = new System.Windows.Forms.TabPage();
            this.UART_Refresh_Button = new System.Windows.Forms.Button();
            this.SerialCancel_Button = new System.Windows.Forms.Button();
            this.SerialOK_Button = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.UART_FlowCtrl_ComboBox = new System.Windows.Forms.ComboBox();
            this.UART_StopBit_ComboBox = new System.Windows.Forms.ComboBox();
            this.UART_Parity_ComboBox = new System.Windows.Forms.ComboBox();
            this.UART_DataBits_ComboBox = new System.Windows.Forms.ComboBox();
            this.UART_BaudRate_ComboBox = new System.Windows.Forms.ComboBox();
            this.UART_Port_ComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SPI_TabPage = new System.Windows.Forms.TabPage();
            this.SPI_Refresh_Button = new System.Windows.Forms.Button();
            this.label16 = new System.Windows.Forms.Label();
            this.SPI_Device_ComboBox = new System.Windows.Forms.ComboBox();
            this.SPI_Mode_ComboBox = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.SPI_BitFirst_ComboBox = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.SPI_Clock_TextBox = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.SPI_Cancel_Button = new System.Windows.Forms.Button();
            this.SPI_OK_Button = new System.Windows.Forms.Button();
            this.I2C_TabPage = new System.Windows.Forms.TabPage();
            this.I2C_Refresh_Button = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.I2C_Device_ComboBox = new System.Windows.Forms.ComboBox();
            this.I2C_SlaveAddr_TextBox = new System.Windows.Forms.TextBox();
            this.label24 = new System.Windows.Forms.Label();
            this.I2C_Clock_TextBox = new System.Windows.Forms.TextBox();
            this.label26 = new System.Windows.Forms.Label();
            this.I2C_Cancel_Button = new System.Windows.Forms.Button();
            this.I2C_OK_Button = new System.Windows.Forms.Button();
            this.Comn_TabControl.SuspendLayout();
            this.Serial_TabPage.SuspendLayout();
            this.SPI_TabPage.SuspendLayout();
            this.I2C_TabPage.SuspendLayout();
            base.SuspendLayout();
            this.Comn_TabControl.Controls.Add(this.Serial_TabPage);
            this.Comn_TabControl.Controls.Add(this.SPI_TabPage);
            this.Comn_TabControl.Controls.Add(this.I2C_TabPage);
            this.Comn_TabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Comn_TabControl.Location = new System.Drawing.Point(0, 0);
            this.Comn_TabControl.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Comn_TabControl.Name = "Comn_TabControl";
            this.Comn_TabControl.SelectedIndex = 0;
            this.Comn_TabControl.Size = new System.Drawing.Size(582, 303);
            this.Comn_TabControl.TabIndex = 0;
            this.Comn_TabControl.Selected += new System.Windows.Forms.TabControlEventHandler(Comn_TabControl_Selected);
            this.Serial_TabPage.Controls.Add(this.UART_Refresh_Button);
            this.Serial_TabPage.Controls.Add(this.SerialCancel_Button);
            this.Serial_TabPage.Controls.Add(this.SerialOK_Button);
            this.Serial_TabPage.Controls.Add(this.label6);
            this.Serial_TabPage.Controls.Add(this.label5);
            this.Serial_TabPage.Controls.Add(this.label4);
            this.Serial_TabPage.Controls.Add(this.label3);
            this.Serial_TabPage.Controls.Add(this.label2);
            this.Serial_TabPage.Controls.Add(this.UART_FlowCtrl_ComboBox);
            this.Serial_TabPage.Controls.Add(this.UART_StopBit_ComboBox);
            this.Serial_TabPage.Controls.Add(this.UART_Parity_ComboBox);
            this.Serial_TabPage.Controls.Add(this.UART_DataBits_ComboBox);
            this.Serial_TabPage.Controls.Add(this.UART_BaudRate_ComboBox);
            this.Serial_TabPage.Controls.Add(this.UART_Port_ComboBox);
            this.Serial_TabPage.Controls.Add(this.label1);
            this.Serial_TabPage.Location = new System.Drawing.Point(4, 25);
            this.Serial_TabPage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Serial_TabPage.Name = "Serial_TabPage";
            this.Serial_TabPage.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Serial_TabPage.Size = new System.Drawing.Size(574, 274);
            this.Serial_TabPage.TabIndex = 0;
            this.Serial_TabPage.Text = "Serial";
            this.Serial_TabPage.UseVisualStyleBackColor = true;
            this.UART_Refresh_Button.Location = new System.Drawing.Point(326, 22);
            this.UART_Refresh_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_Refresh_Button.Name = "UART_Refresh_Button";
            this.UART_Refresh_Button.Size = new System.Drawing.Size(75, 30);
            this.UART_Refresh_Button.TabIndex = 58;
            this.UART_Refresh_Button.Text = "Refresh";
            this.UART_Refresh_Button.UseVisualStyleBackColor = true;
            this.UART_Refresh_Button.Click += new System.EventHandler(UART_Refresh_Button_Click);
            this.SerialCancel_Button.Location = new System.Drawing.Point(435, 75);
            this.SerialCancel_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SerialCancel_Button.Name = "SerialCancel_Button";
            this.SerialCancel_Button.Size = new System.Drawing.Size(130, 35);
            this.SerialCancel_Button.TabIndex = 13;
            this.SerialCancel_Button.Text = "Cancel";
            this.SerialCancel_Button.UseVisualStyleBackColor = true;
            this.SerialCancel_Button.Click += new System.EventHandler(UART_Cancel_Button_Click);
            this.SerialOK_Button.Location = new System.Drawing.Point(435, 25);
            this.SerialOK_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SerialOK_Button.Name = "SerialOK_Button";
            this.SerialOK_Button.Size = new System.Drawing.Size(130, 35);
            this.SerialOK_Button.TabIndex = 12;
            this.SerialOK_Button.Text = "OK";
            this.SerialOK_Button.UseVisualStyleBackColor = true;
            this.SerialOK_Button.Click += new System.EventHandler(UART_OK_Button_Click);
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(19, 216);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(99, 15);
            this.label6.TabIndex = 11;
            this.label6.Text = "Flow Control :";
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 179);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 15);
            this.label5.TabIndex = 10;
            this.label5.Text = "Stop Bit :";
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 141);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(54, 15);
            this.label4.TabIndex = 9;
            this.label4.Text = "Parity :";
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(19, 106);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 15);
            this.label3.TabIndex = 8;
            this.label3.Text = "Data Bits :";
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 69);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 15);
            this.label2.TabIndex = 7;
            this.label2.Text = "Baud Rate :";
            this.UART_FlowCtrl_ComboBox.FormattingEnabled = true;
            this.UART_FlowCtrl_ComboBox.Location = new System.Drawing.Point(150, 212);
            this.UART_FlowCtrl_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_FlowCtrl_ComboBox.Name = "UART_FlowCtrl_ComboBox";
            this.UART_FlowCtrl_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_FlowCtrl_ComboBox.TabIndex = 6;
            this.UART_StopBit_ComboBox.FormattingEnabled = true;
            this.UART_StopBit_ComboBox.Location = new System.Drawing.Point(150, 175);
            this.UART_StopBit_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_StopBit_ComboBox.Name = "UART_StopBit_ComboBox";
            this.UART_StopBit_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_StopBit_ComboBox.TabIndex = 5;
            this.UART_Parity_ComboBox.FormattingEnabled = true;
            this.UART_Parity_ComboBox.Location = new System.Drawing.Point(150, 138);
            this.UART_Parity_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_Parity_ComboBox.Name = "UART_Parity_ComboBox";
            this.UART_Parity_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_Parity_ComboBox.TabIndex = 4;
            this.UART_DataBits_ComboBox.FormattingEnabled = true;
            this.UART_DataBits_ComboBox.Location = new System.Drawing.Point(150, 100);
            this.UART_DataBits_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_DataBits_ComboBox.Name = "UART_DataBits_ComboBox";
            this.UART_DataBits_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_DataBits_ComboBox.TabIndex = 3;
            this.UART_BaudRate_ComboBox.FormattingEnabled = true;
            this.UART_BaudRate_ComboBox.Location = new System.Drawing.Point(150, 62);
            this.UART_BaudRate_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_BaudRate_ComboBox.Name = "UART_BaudRate_ComboBox";
            this.UART_BaudRate_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_BaudRate_ComboBox.TabIndex = 2;
            this.UART_Port_ComboBox.FormattingEnabled = true;
            this.UART_Port_ComboBox.Location = new System.Drawing.Point(150, 25);
            this.UART_Port_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.UART_Port_ComboBox.Name = "UART_Port_ComboBox";
            this.UART_Port_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.UART_Port_ComboBox.TabIndex = 1;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(19, 31);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Port Number :";
            this.SPI_TabPage.Controls.Add(this.SPI_Refresh_Button);
            this.SPI_TabPage.Controls.Add(this.label16);
            this.SPI_TabPage.Controls.Add(this.SPI_Device_ComboBox);
            this.SPI_TabPage.Controls.Add(this.SPI_Mode_ComboBox);
            this.SPI_TabPage.Controls.Add(this.label10);
            this.SPI_TabPage.Controls.Add(this.SPI_BitFirst_ComboBox);
            this.SPI_TabPage.Controls.Add(this.label8);
            this.SPI_TabPage.Controls.Add(this.SPI_Clock_TextBox);
            this.SPI_TabPage.Controls.Add(this.label7);
            this.SPI_TabPage.Controls.Add(this.SPI_Cancel_Button);
            this.SPI_TabPage.Controls.Add(this.SPI_OK_Button);
            this.SPI_TabPage.Location = new System.Drawing.Point(4, 25);
            this.SPI_TabPage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_TabPage.Name = "SPI_TabPage";
            this.SPI_TabPage.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_TabPage.Size = new System.Drawing.Size(574, 274);
            this.SPI_TabPage.TabIndex = 1;
            this.SPI_TabPage.Text = "SPI";
            this.SPI_TabPage.UseVisualStyleBackColor = true;
            this.SPI_Refresh_Button.Location = new System.Drawing.Point(326, 22);
            this.SPI_Refresh_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_Refresh_Button.Name = "SPI_Refresh_Button";
            this.SPI_Refresh_Button.Size = new System.Drawing.Size(75, 30);
            this.SPI_Refresh_Button.TabIndex = 36;
            this.SPI_Refresh_Button.Text = "Refresh";
            this.SPI_Refresh_Button.UseVisualStyleBackColor = true;
            this.SPI_Refresh_Button.Click += new System.EventHandler(SPI_Refresh_Button_Click);
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(19, 31);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(90, 15);
            this.label16.TabIndex = 35;
            this.label16.Text = "Device List :";
            this.SPI_Device_ComboBox.FormattingEnabled = true;
            this.SPI_Device_ComboBox.Location = new System.Drawing.Point(150, 25);
            this.SPI_Device_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_Device_ComboBox.Name = "SPI_Device_ComboBox";
            this.SPI_Device_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.SPI_Device_ComboBox.TabIndex = 34;
            this.SPI_Mode_ComboBox.FormattingEnabled = true;
            this.SPI_Mode_ComboBox.Location = new System.Drawing.Point(150, 100);
            this.SPI_Mode_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_Mode_ComboBox.Name = "SPI_Mode_ComboBox";
            this.SPI_Mode_ComboBox.Size = new System.Drawing.Size(168, 23);
            this.SPI_Mode_ComboBox.TabIndex = 23;
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(19, 106);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(82, 15);
            this.label10.TabIndex = 22;
            this.label10.Text = "SPI Mode :";
            this.SPI_BitFirst_ComboBox.FormattingEnabled = true;
            this.SPI_BitFirst_ComboBox.Location = new System.Drawing.Point(150, 138);
            this.SPI_BitFirst_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_BitFirst_ComboBox.Name = "SPI_BitFirst_ComboBox";
            this.SPI_BitFirst_ComboBox.Size = new System.Drawing.Size(99, 23);
            this.SPI_BitFirst_ComboBox.TabIndex = 19;
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(19, 144);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(106, 15);
            this.label8.TabIndex = 18;
            this.label8.Text = "Significant Bit :";
            this.SPI_Clock_TextBox.Location = new System.Drawing.Point(150, 62);
            this.SPI_Clock_TextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_Clock_TextBox.Name = "SPI_Clock_TextBox";
            this.SPI_Clock_TextBox.Size = new System.Drawing.Size(99, 25);
            this.SPI_Clock_TextBox.TabIndex = 17;
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(19, 69);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(95, 15);
            this.label7.TabIndex = 16;
            this.label7.Text = "SCLK (kHz) :";
            this.SPI_Cancel_Button.Location = new System.Drawing.Point(435, 75);
            this.SPI_Cancel_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_Cancel_Button.Name = "SPI_Cancel_Button";
            this.SPI_Cancel_Button.Size = new System.Drawing.Size(130, 35);
            this.SPI_Cancel_Button.TabIndex = 15;
            this.SPI_Cancel_Button.Text = "Cancel";
            this.SPI_Cancel_Button.UseVisualStyleBackColor = true;
            this.SPI_Cancel_Button.Click += new System.EventHandler(SPI_Cancel_Button_Click);
            this.SPI_OK_Button.Location = new System.Drawing.Point(435, 25);
            this.SPI_OK_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.SPI_OK_Button.Name = "SPI_OK_Button";
            this.SPI_OK_Button.Size = new System.Drawing.Size(130, 35);
            this.SPI_OK_Button.TabIndex = 14;
            this.SPI_OK_Button.Text = "OK";
            this.SPI_OK_Button.UseVisualStyleBackColor = true;
            this.SPI_OK_Button.Click += new System.EventHandler(SPI_OK_Button_Click);
            this.I2C_TabPage.Controls.Add(this.I2C_Refresh_Button);
            this.I2C_TabPage.Controls.Add(this.label17);
            this.I2C_TabPage.Controls.Add(this.I2C_Device_ComboBox);
            this.I2C_TabPage.Controls.Add(this.I2C_SlaveAddr_TextBox);
            this.I2C_TabPage.Controls.Add(this.label24);
            this.I2C_TabPage.Controls.Add(this.I2C_Clock_TextBox);
            this.I2C_TabPage.Controls.Add(this.label26);
            this.I2C_TabPage.Controls.Add(this.I2C_Cancel_Button);
            this.I2C_TabPage.Controls.Add(this.I2C_OK_Button);
            this.I2C_TabPage.Location = new System.Drawing.Point(4, 25);
            this.I2C_TabPage.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_TabPage.Name = "I2C_TabPage";
            this.I2C_TabPage.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_TabPage.Size = new System.Drawing.Size(574, 274);
            this.I2C_TabPage.TabIndex = 2;
            this.I2C_TabPage.Text = "I2C";
            this.I2C_TabPage.UseVisualStyleBackColor = true;
            this.I2C_Refresh_Button.Location = new System.Drawing.Point(326, 22);
            this.I2C_Refresh_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_Refresh_Button.Name = "I2C_Refresh_Button";
            this.I2C_Refresh_Button.Size = new System.Drawing.Size(75, 30);
            this.I2C_Refresh_Button.TabIndex = 57;
            this.I2C_Refresh_Button.Text = "Refresh";
            this.I2C_Refresh_Button.UseVisualStyleBackColor = true;
            this.I2C_Refresh_Button.Click += new System.EventHandler(I2C_Refresh_Button_Click);
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(19, 31);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(90, 15);
            this.label17.TabIndex = 56;
            this.label17.Text = "Device List :";
            this.I2C_Device_ComboBox.FormattingEnabled = true;
            this.I2C_Device_ComboBox.Location = new System.Drawing.Point(169, 25);
            this.I2C_Device_ComboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_Device_ComboBox.Name = "I2C_Device_ComboBox";
            this.I2C_Device_ComboBox.Size = new System.Drawing.Size(149, 23);
            this.I2C_Device_ComboBox.TabIndex = 55;
            this.I2C_SlaveAddr_TextBox.Location = new System.Drawing.Point(169, 100);
            this.I2C_SlaveAddr_TextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_SlaveAddr_TextBox.Name = "I2C_SlaveAddr_TextBox";
            this.I2C_SlaveAddr_TextBox.Size = new System.Drawing.Size(99, 25);
            this.I2C_SlaveAddr_TextBox.TabIndex = 43;
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(19, 106);
            this.label24.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(132, 15);
            this.label24.TabIndex = 42;
            this.label24.Text = "Slave addr (HEX) :";
            this.I2C_Clock_TextBox.Location = new System.Drawing.Point(169, 62);
            this.I2C_Clock_TextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_Clock_TextBox.Name = "I2C_Clock_TextBox";
            this.I2C_Clock_TextBox.Size = new System.Drawing.Size(99, 25);
            this.I2C_Clock_TextBox.TabIndex = 39;
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(19, 69);
            this.label26.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(129, 15);
            this.label26.TabIndex = 38;
            this.label26.Text = "SCL Clock (kHz) :";
            this.I2C_Cancel_Button.Location = new System.Drawing.Point(435, 75);
            this.I2C_Cancel_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_Cancel_Button.Name = "I2C_Cancel_Button";
            this.I2C_Cancel_Button.Size = new System.Drawing.Size(130, 35);
            this.I2C_Cancel_Button.TabIndex = 37;
            this.I2C_Cancel_Button.Text = "Cancel";
            this.I2C_Cancel_Button.UseVisualStyleBackColor = true;
            this.I2C_Cancel_Button.Click += new System.EventHandler(I2C_Cancel_Button_Click);
            this.I2C_OK_Button.Location = new System.Drawing.Point(435, 25);
            this.I2C_OK_Button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.I2C_OK_Button.Name = "I2C_OK_Button";
            this.I2C_OK_Button.Size = new System.Drawing.Size(130, 35);
            this.I2C_OK_Button.TabIndex = 36;
            this.I2C_OK_Button.Text = "OK";
            this.I2C_OK_Button.UseVisualStyleBackColor = true;
            this.I2C_OK_Button.Click += new System.EventHandler(I2C_OK_Button_Click);
            base.AutoScaleDimensions = new System.Drawing.SizeF(120f, 120f);
            base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            base.ClientSize = new System.Drawing.Size(582, 303);
            base.Controls.Add(this.Comn_TabControl);
            base.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            base.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            base.Name = "WireComnForm";
            this.Text = "Wire Communication Setup";
            base.FormClosing += new System.Windows.Forms.FormClosingEventHandler(WireComnForm_FormClosing);
            base.Load += new System.EventHandler(WireComnForm_Load);
            this.Comn_TabControl.ResumeLayout(false);
            this.Serial_TabPage.ResumeLayout(false);
            this.Serial_TabPage.PerformLayout();
            this.SPI_TabPage.ResumeLayout(false);
            this.SPI_TabPage.PerformLayout();
            this.I2C_TabPage.ResumeLayout(false);
            this.I2C_TabPage.PerformLayout();
            base.ResumeLayout(false);
        }
    }
}
