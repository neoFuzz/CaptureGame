using DirectShowLib;
using System;

namespace ConsTest
{
    class Program
    {
        private static IGraphBuilder m_objFilterGraph;
        private static IBasicAudio m_objBasicAudio;
        private static IMediaControl m_objMediaControl;

        static void Main(string[] args)
        {
            int count = 0, readA = 0, readI = 0;

            DsDevice[] devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("***** Renderers *****");
            Console.ResetColor();
            foreach (DsDevice dsd in devices)
            {
                Console.WriteLine("[" + count + "] " + dsd.Name);
                count++;
            }
            count--; // Subtract one to make it zero-indexed
            readA = getDeviceNumber(count);
            Console.WriteLine("Selected Device: " + devices[readA].Name);

            count = 0;
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioInputDevice);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("***** Input Devices *****");
            Console.ResetColor();
            foreach (DsDevice dsd in devices)
            {
                Console.WriteLine("[" + count + "] " + dsd.Name);
                count++;
            }
            count--; // Subtract one to make it zero-indexed
            readI = getDeviceNumber(count);
            Console.WriteLine("Selected Device: " + devices[readI].Name);
            Connect_up(readI, readA);

            setUpSound();

            m_objMediaControl.Run();
            // Wait to end
            Console.ReadLine();
            m_objMediaControl.Stop();
            GC.Collect();
        }

        private static void setUpSound()
        {

        }

        private static void Connect_up(int dInput, int dOutput)
        {
            object source = null, inputD = null;
            DsDevice[] devices;
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            DsDevice device = devices[dOutput];
            Guid iid = typeof(IBaseFilter).GUID;
            device.Mon.BindToObject(null, null, ref iid, out source);

            m_objFilterGraph = (IGraphBuilder)new FilterGraph();
            m_objFilterGraph.AddFilter((IBaseFilter)source, "Audio Input pin (rendered)");

            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioInputDevice);
            device = devices[dInput];
            // iid change?
            device.Mon.BindToObject(null, null, ref iid, out inputD);

            m_objFilterGraph.AddFilter((IBaseFilter)inputD, "Capture");

            int result;
            IEnumPins pInputPin = null, pOutputPin = null;// Pin enumeration
            IPin pIn = null, pOut = null;// Pins
            try
            {
                IBaseFilter newI = (IBaseFilter)inputD;
                result = newI.EnumPins(out pInputPin);// Enumerate the pin
                if (result.Equals(0))
                {
                    // Get hold of the pin as seen in GraphEdit
                    newI.FindPin("Capture", out pIn);
                }
                IBaseFilter ibfO = (IBaseFilter)source;
                ibfO.EnumPins(out pOutputPin);//Enumerate the pin

                ibfO.FindPin("Audio Input pin (rendered)", out pOut);
                try
                {
                    pIn.Connect(pOut, null);  // Connect the input pin to output pin
                }
                catch (Exception ex)
                { Console.WriteLine(ex.Message); }
            }
            catch (Exception ex)
            { Console.WriteLine(ex.Message); }

            m_objBasicAudio = m_objFilterGraph as IBasicAudio;
            m_objMediaControl = m_objFilterGraph as IMediaControl;

        }
        /// <summary>
        /// Prompt for device number and take valid input from user
        /// </summary>
        /// <param name="countOfDev">Zero-indexed total number of devices to choose from</param>
        /// <returns>integer based on user input</returns>
        private static int getDeviceNumber(int countOfDev)
        {
            int rDevice = 0;
            bool testVal = false;
            Console.Write("Device [0-{0}]: ", countOfDev);
            while (testVal == false)
            {
                try { rDevice = Convert.ToInt32(Console.ReadLine()); }
                catch (Exception ex) { }
                if (rDevice <= countOfDev)
                {
                    testVal = true;
                    break;
                }
                Console.Write("Device [0-{0}]: ", countOfDev);
            }
            return rDevice;
        }
        private void playAudioToDevice(string fName, int devIndex)
        {
            object source = null;
            DsDevice[] devices;
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            DsDevice device = (DsDevice)devices[devIndex];
            Guid iid = typeof(IBaseFilter).GUID;
            device.Mon.BindToObject(null, null, ref iid, out source);

            m_objFilterGraph = (IGraphBuilder)new FilterGraph();
            m_objFilterGraph.AddFilter((IBaseFilter)source, "Audio Render");
            m_objFilterGraph.RenderFile(fName, "");

            m_objBasicAudio = m_objFilterGraph as IBasicAudio;
            m_objMediaControl = m_objFilterGraph as IMediaControl;

            m_objMediaControl.Run();
        }
    }
}
