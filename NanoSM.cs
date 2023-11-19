//tabs=4
// --------------------------------------------------------------------------------
// 
// ASCOM SafetyMonitor driver for NanoWeather Arduino project
//
// Description:	Part of a NanoEvery project to read sky, environment data for general use
//
// Implements:	ASCOM SafetyMonitor interface version: 1.5
// Author:		(CJW) chris@digitalastrophotography.co.uk>
//
// Edit Log:
//
// Date			Who	Vers	Description
// -----------	---	-----	-------------------------------------------------------
// 25-Oct-2020	CJW	1.0.0	Initial edit, created from ASCOM driver template
// 26-Oct-2020  CJW 1.1.0   Corrected string comparison - changed parsing
// 30-Oct-2020  CJW 1.2     Serial.dispose added when closing connection
// 30-Oct-2020  CJW 1.3     Out of bound thresholds revert to defaults
// 5 Nov 2020   CJW 1.4     changed controls on chooser and used device profile to store
// 23 Aug 2021  CJW 1.5     added serial controls for buzzer and reset ... more to come
// --------------------------------------------------------------------------------
//


// This is used to define code in the template that is specific to one class implementation
// unused code can be deleted and this definition removed.
#define SafetyMonitor

using ASCOM.Astrometry.AstroUtils;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;

namespace ASCOM.NanoSM
{
    //
    // Your driver's DeviceID is ASCOM.NanoSM.SafetyMonitor
    //
    // The Guid attribute sets the CLSID for ASCOM.NanoSM.SafetyMonitor
    // The ClassInterface/None attribute prevents an empty interface called
    // _NanoSM from being created and used as the [default] interface

    /// <summary>
    /// ASCOM SafetyMonitor Driver for NanoSM.
    /// </summary>
    enum MessageIndex { tempC, humdity, pressure, skytemp, ambient, skymag, rainratio }
    [Guid("76a5a579-579d-4e93-987d-79a65240c62e")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SafetyMonitor : ISafetyMonitor
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal static string driverID = "ASCOM.NanoSM.SafetyMonitor";
        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "NanoSM SafetyMonitor";

        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>

        internal static double RainRatioThreshold, SkyTempThreshold, HumidityThreshold;  // set up by ASCOM chooser
        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the current device configuration
        internal string buffer = "";
        internal bool startCharRx;  // indicates start of message detected
        internal char endChar = '#';  // end character of Arduino response
        internal char startChar = '$';  // start character of Arduino response
        //internal char[] delimeter = new char[] {','};  // delimiter between Arduino response values
        internal char[] delimeter = { ',' };
        internal string arduinoMessage = ""; // string that builds up received message
        internal bool dataRx = false;  // indicates that data is available to read

        //  these values may be populated by setup chooser screen in the future
        string[] configure = new string[3]; // for thresholds
        internal double K1 = 33;   // skytemp modification with air temp (K1/100)
        internal double K3 = 4;    // offset for skytemp (-K3*K1/100 )
        private string[] NanoStatus = new string[7];  //  "$tempC, humidity, pressure, skyrawC, ambientC, skymag, rainratio #"

        /// </summary>
        private bool connectedState;

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Variable to hold the trace logger object (creates a diagnostic log file with information that you specify)
        /// </summary>
        internal TraceLogger tl;
        private SerialPort Serial;  // my serial port instance of ASCOM serial port

        /// <summary>
        /// Initializes a new instance of the <see cref="NanoSM"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public SafetyMonitor()
        {
            tl = new TraceLogger("", "NanoSM");
            ReadProfile(); // Read device configuration from the ASCOM Profile store

            tl.LogMessage("SafetyMonitor", "Starting initialisation");

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro-utilities object
            Serial = new SerialPort();  // standard .net serial port
            tl.LogMessage("SafetyMonitor", "Completed initialisation");
        }
        // openArduino initilises serial port and set up an event handler to suck in characters
        // it runs in the background, Arduino broadcasts every 2 seconds
        private bool openArduino()
        {
            Serial.BaudRate = 9600;
            Serial.PortName = comPort;
            Serial.Parity = Parity.None;
            Serial.DataBits = 8;
            Serial.Handshake = System.IO.Ports.Handshake.None;
            Serial.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(receiveData);
            Serial.ReceivedBytesThreshold = 1;
            Serial.DtrEnable = true;   // needed to prompt Arduino to reset and pump out of USB serial
            try
            {
                Serial.Open();              // open port
                Serial.DiscardInBuffer();   // and clear it out just in case
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }


        // receiveData is based on a code fragment suggested by Per and reads characters as they arrive
        // it decodes the messages, looking for framing characters and then splits the CSV string into
        // component parts to represent the status flags from the Arduino 
        private void receiveData(object sender, SerialDataReceivedEventArgs e)
        {
            if (e.EventType == System.IO.Ports.SerialData.Chars)
            {
                while (Serial.BytesToRead > 0)
                {
                    char c = (char)Serial.ReadChar();  // wait for start character
                    if (!startCharRx)
                    {
                        if (c == startChar)  // and then initialise the message
                        {
                            startCharRx = true;
                            buffer = "";  // clear buffer
                        }
                    }
                    else
                    {
                        if (c == endChar)
                        {
                            arduinoMessage = buffer;  // transfer the buffer to the message and clear the buffer
                            buffer = "";
                            startCharRx = false;
                            if (arduinoMessage.Length <= 39) // check the message length is OK  38 is longest, below zero and high pressure
                            {
                                dataRx = true; // tell the world that data is available
                                NanoStatus = arduinoMessage.Split(delimeter);  // temp, humidity, pressure, sky, ambient, magnitude, rain
                                tl.LogMessage("communications", arduinoMessage);
                            }
                            else  // message was corrupted
                            {
                                dataRx = false;
                                tl.LogMessage("communications", "corrupted message length");
                                arduinoMessage = "";
                            }
                        }
                        else
                        {
                            buffer += c;  // build up message string in buffer
                        }
                    }
                }
            }
        }

        private void Diagnostic()  // call this to check received string
        {
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.rainratio]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.skytemp]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.tempC]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.humdity]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.skytemp]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.tempC]);
            tl.LogMessage("values", NanoStatus[(int)MessageIndex.humdity]);
        }


        // calculates the safety status - later version will be able to change thresholds
        private Boolean SafetyStatus()
        {
            if (dataRx)
            {
                double rain, skytemp, airtemp, humidity;
                if (double.TryParse(NanoStatus[(int)MessageIndex.humdity], out humidity))
                    if (humidity > HumidityThreshold) return false; // if extremely humid
                if (double.TryParse(NanoStatus[(int)MessageIndex.rainratio], out rain))
                    if (rain > RainRatioThreshold) return false;
                if (double.TryParse(NanoStatus[(int)MessageIndex.skytemp], out skytemp))
                    if (double.TryParse(NanoStatus[(int)MessageIndex.tempC], out airtemp))
                    {
                        skytemp = skytemp - (airtemp + K3) * K1 / 100; // modified airtemp
                        if (skytemp > SkyTempThreshold) return false; // if temperature is not low enough, overcast 
                        return true;
                    }
                    else
                    {
                        tl.LogMessage("communications", "summit wrong");
                        return false;
                    }
            }
            tl.LogMessage("communications", "no data ");
            return false;
        }

        //
        // PUBLIC COM INTERFACE ISafetyMonitor IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (NanoSMSetup F = new NanoSMSetup(tl))
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                tl.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            CheckConnected("CommandBlind");
            CommandString(command, raw);  // don't need the return string
            return;
        }

        public bool CommandBool(string command, bool raw)
        {
            tl.LogMessage("CommandBool", "Not implemented");
            throw new ASCOM.MethodNotImplementedException("CommandBool");
        }

        public string CommandString(string command, bool raw)
        {
            CheckConnected("CommandString");
            // this is the customised I/O to the serial port, used by all commands
            // status commands are interpreted from cache variables, and commands
            // are issued
            try
            {
                if (!raw) // if command uses delimiter
                {
                    tl.LogMessage("attempting commandstring", command);
                    if (IsConnected)  // only if connected  - try and avoid comms error
                    {
                        if (command == "reset" ||  command == "buzz on"|| command == "buzz off")
                        {
                            Serial.Write(command + "#"); 
                            return ("1");
                        }
                    }
                    throw new ASCOM.NotConnectedException("com port not connected");
                }
                else  // do nothing if the command is not using delimiter
                {
                    tl.LogMessage("commandstring ", "Not implemented without # terminator");
                    throw new ASCOM.MethodNotImplementedException("CommandString");
                }
            }
            catch (Exception)  // better luck next time :)
            {
                System.Windows.Forms.MessageBox.Show("Timed out, press OK to recover");
                return ("comms error");
            }
        }

        public void Dispose()
        {
            // Clean up the trace logger and util objects
            tl.Enabled = false;
            tl.Dispose();
            tl = null;
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
            Serial.Dispose();
            Serial = null;
        }

        public bool Connected
        {
            get
            {
                LogMessage("Connected", "Get {0}", IsConnected);
                return IsConnected;
            }
            set
            {
                tl.LogMessage("Connected", "Set {0}", value);
                if (value == IsConnected)
                    return;

                if (value)
                {

                    LogMessage("Connected Set", "Connecting to port {0}", comPort);
                    if (!openArduino())
                    {
                        Serial.Dispose();
                        LogMessage("connection", "Problem with connecting", comPort);
                    }
                    else
                    {
                        connectedState = true;
                        for (int i = 0; i < 7; i++) NanoStatus[i] = "1.0";  // safe default values
                    }
                }
                else
                {
                    connectedState = false;
                    LogMessage("Connected Set", "Disconnecting from port {0}", comPort);
                    Serial.Dispose();  //disconnect to serial
                    Serial = null;
                }
            }
        }

        public string Description
        {
            get
            {
                tl.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverInfo = "Nano Environment monitor" + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                tl.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "1");
                return Convert.ToInt16("1");
            }
        }

        public string Name
        {
            get
            {
                string name = "NanoEvery Cloud";
                tl.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region ISafetyMonitor Implementation
        public bool IsSafe
        {
            get
            {
                if (Connected) return SafetyStatus();
                else return false;  //  default before connection is false 
            }
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "SafetyMonitor";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // check the actual serial connection (checks for unplugged)
                connectedState = Serial.IsOpen;
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "SafetyMonitor";
                tl.Enabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
                HumidityThreshold = double.Parse(driverProfile.GetValue(driverID, "humidLimit", string.Empty, "90"));
                SkyTempThreshold = double.Parse(driverProfile.GetValue(driverID, "skyThreshold", string.Empty, "0"));
                RainRatioThreshold = double.Parse(driverProfile.GetValue(driverID, "rainThreshold", string.Empty, "1.2"));
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "SafetyMonitor";
                driverProfile.WriteValue(driverID, traceStateProfileName, tl.Enabled.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
                driverProfile.WriteValue(driverID, "humidLimit", HumidityThreshold.ToString());
                driverProfile.WriteValue(driverID, "skyThreshold", SkyTempThreshold.ToString());
                driverProfile.WriteValue(driverID, "rainThreshold", RainRatioThreshold.ToString());
            }
        }

        /// <summary>
        /// Log helper function that takes formatted strings and arguments
        /// </summary>
        /// <param name="identifier"></param>
        /// <param name="message"></param>
        /// <param name="args"></param>
        internal void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            tl.LogMessage(identifier, msg);
        }
        #endregion
    }
}
