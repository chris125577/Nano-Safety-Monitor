using ASCOM.Utilities;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ASCOM.NanoSM;

namespace ASCOM.NanoSM
{
    [ComVisible(false)]					// Form not registered for COM!
    public partial class NanoSMSetup : Form
    {
        TraceLogger tl; // Holder for a reference to the driver's trace logger

        public NanoSMSetup(TraceLogger tlDriver)
        {
            InitializeComponent();

            // Save the provided trace logger for use within the setup dialogue
            tl = tlDriver;

            // Initialise current values of user settings from the ASCOM Profile
            InitUI();
        }

        private void cmdOK_Click(object sender, EventArgs e) // OK button event handler
        {
            SafetyMonitor.comPort = (string)comboBoxComPort.SelectedItem;
            tl.Enabled = chkTrace.Checked;
            SafetyMonitor.HumidityThreshold = (double)inputHumid.Value;
            SafetyMonitor.RainRatioThreshold = (double)inputRain.Value;
            SafetyMonitor.SkyTempThreshold = (double)inputCloud.Value;
        }

        private void cmdCancel_Click(object sender, EventArgs e) // Cancel button event handler
        {
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e) // Click on ASCOM logo event handler
        {
            try
            {
                System.Diagnostics.Process.Start("http://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (System.Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }

        private void InitUI()
        {
            chkTrace.Checked = tl.Enabled;
            // set the list of com ports to those that are currently available
            comboBoxComPort.Items.Clear();
            comboBoxComPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());      // use System.IO because it's static
            inputCloud.Value = (decimal)SafetyMonitor.SkyTempThreshold;
            inputHumid.Value = (decimal)SafetyMonitor.HumidityThreshold;
            inputRain.Value = (decimal)SafetyMonitor.RainRatioThreshold;
            // select the current port if possible
            if (comboBoxComPort.Items.Contains(SafetyMonitor.comPort))
            {
                comboBoxComPort.SelectedItem = SafetyMonitor.comPort;
            }
        }
    }
}