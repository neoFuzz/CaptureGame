using CaptureGame.Properties;
using DirectX.Capture;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DirectShowLib;
using TunerInputType = DirectX.Capture.TunerInputType;

namespace CaptureGame
{
    public class FrmCaptureGame : System.Windows.Forms.Form
    {
        private Capture capture = null;
        private Filters filters = new Filters();
        private MainMenu mainMenu;
        private Panel panelVideo;
        private MenuStrip msMainMenu;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem devicesToolStripMenuItem;
        private ToolStripMenuItem videoDevicesToolStripMenuItem;
        private ToolStripMenuItem testToolStripMenuItem;
        private ToolStripMenuItem audioDevicesToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripMenuItem videoCompressorsToolStripMenuItem;
        private ToolStripMenuItem audioCompressorsToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripMenuItem scanDevicesToolStripMenuItem;
        private ToolStripMenuItem optionsToolStripMenuItem;
        private ToolStripMenuItem videoSourcesToolStripMenuItem;
        private ToolStripMenuItem videoFrameSizeToolStripMenuItem;
        private ToolStripMenuItem videoFrameRateToolStripMenuItem;
        private ToolStripMenuItem videoCapabilitiesToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripMenuItem audioSourcesToolStripMenuItem;
        private ToolStripMenuItem audioChannelsToolStripMenuItem;
        private ToolStripMenuItem audioSamplingRateToolStripMenuItem;
        private ToolStripMenuItem audioSampleSizeToolStripMenuItem;
        private ToolStripMenuItem audioCapabilitiesToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripMenuItem tVTunerChannelToolStripMenuItem;
        private ToolStripMenuItem tVTunerInputTypeToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator5;
        private ToolStripMenuItem propertyPagesToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator6;
        private ToolStripMenuItem previewToolStripMenuItem;
        private ToolStripMenuItem testAToolStripMenuItem;
        private IContainer components;
        private ToolStripMenuItem audioOutputToolStripMenuItem;
        private ToolStripMenuItem breakerToolStripMenuItem;
        /// <summary>
        /// TODO: find this and add audio
        /// </summary>
        private static IGraphBuilder m_objFilterGraph = null;
        private static IBasicAudio m_objBasicAudio = null;
        private static IMediaControl m_objMediaControl = null;
        private ToolStripMenuItem GameNameToolStripMenuItem;
        private ToolStripSeparator toolStripSeparator7;
        private DsDevice[] devices;

        public FrmCaptureGame()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            // Start with the first video/audio devices
            // Don't do this in the Release build in case the
            // first devices cause problems.
#if DEBUG
            capture = new Capture(filters.VideoInputDevices[0], filters.AudioInputDevices[0]);
            capture.CaptureComplete += new EventHandler(OnCaptureComplete);
#endif
            // Update the main menu
            // Much of the interesting work of this sample occurs here
            try {
                updateMenu();
                m_objFilterGraph = (IGraphBuilder) new FilterGraph();
                m_objMediaControl = m_objFilterGraph as IMediaControl;
            } catch { }
            Activate();
            this.Text = Settings.Default.GameName;
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
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
            this.components = new System.ComponentModel.Container();
            this.mainMenu = new System.Windows.Forms.MainMenu(this.components);
            this.panelVideo = new System.Windows.Forms.Panel();
            this.msMainMenu = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.GameNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.devicesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.videoDevicesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioDevicesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.testAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioOutputToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.breakerToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.videoCompressorsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioCompressorsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.scanDevicesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.videoSourcesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.videoFrameSizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.videoFrameRateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.videoCapabilitiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.audioSourcesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioChannelsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioSamplingRateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioSampleSizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.audioCapabilitiesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.tVTunerChannelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tVTunerInputTypeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.propertyPagesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.previewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
            this.msMainMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelVideo
            // 
            this.panelVideo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelVideo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelVideo.Location = new System.Drawing.Point(0, 24);
            this.panelVideo.Name = "panelVideo";
            this.panelVideo.Size = new System.Drawing.Size(496, 361);
            this.panelVideo.TabIndex = 6;
            this.panelVideo.Paint += new System.Windows.Forms.PaintEventHandler(this.PanelVideo_Paint);
            this.panelVideo.MouseClick += new System.Windows.Forms.MouseEventHandler(this.PanelVideo_MouseClick);
            // 
            // msMainMenu
            // 
            this.msMainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.devicesToolStripMenuItem,
            this.optionsToolStripMenuItem});
            this.msMainMenu.Location = new System.Drawing.Point(0, 0);
            this.msMainMenu.Name = "msMainMenu";
            this.msMainMenu.Size = new System.Drawing.Size(496, 24);
            this.msMainMenu.TabIndex = 7;
            this.msMainMenu.Text = "msMainMenu";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.GameNameToolStripMenuItem,
            this.toolStripSeparator7,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            this.fileToolStripMenuItem.Click += new System.EventHandler(this.FileToolStripMenuItem_Click);
            // 
            // GameNameToolStripMenuItem
            // 
            this.GameNameToolStripMenuItem.Name = "GameNameToolStripMenuItem";
            this.GameNameToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.GameNameToolStripMenuItem.Text = "Game Name";
            this.GameNameToolStripMenuItem.Click += new System.EventHandler(this.GameNameToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItem_Click);
            // 
            // devicesToolStripMenuItem
            // 
            this.devicesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.videoDevicesToolStripMenuItem,
            this.audioDevicesToolStripMenuItem,
            this.audioOutputToolStripMenuItem,
            this.toolStripSeparator1,
            this.videoCompressorsToolStripMenuItem,
            this.audioCompressorsToolStripMenuItem,
            this.toolStripSeparator2,
            this.scanDevicesToolStripMenuItem});
            this.devicesToolStripMenuItem.Name = "devicesToolStripMenuItem";
            this.devicesToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.devicesToolStripMenuItem.Text = "&Devices";
            // 
            // videoDevicesToolStripMenuItem
            // 
            this.videoDevicesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testToolStripMenuItem});
            this.videoDevicesToolStripMenuItem.Name = "videoDevicesToolStripMenuItem";
            this.videoDevicesToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.videoDevicesToolStripMenuItem.Text = "Video Devices";
            // 
            // testToolStripMenuItem
            // 
            this.testToolStripMenuItem.Name = "testToolStripMenuItem";
            this.testToolStripMenuItem.Size = new System.Drawing.Size(93, 22);
            this.testToolStripMenuItem.Text = "test";
            // 
            // audioDevicesToolStripMenuItem
            // 
            this.audioDevicesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.testAToolStripMenuItem});
            this.audioDevicesToolStripMenuItem.Name = "audioDevicesToolStripMenuItem";
            this.audioDevicesToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.audioDevicesToolStripMenuItem.Text = "Audio Devices";
            // 
            // testAToolStripMenuItem
            // 
            this.testAToolStripMenuItem.Name = "testAToolStripMenuItem";
            this.testAToolStripMenuItem.Size = new System.Drawing.Size(101, 22);
            this.testAToolStripMenuItem.Text = "testA";
            // 
            // audioOutputToolStripMenuItem
            // 
            this.audioOutputToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.breakerToolStripMenuItem});
            this.audioOutputToolStripMenuItem.Name = "audioOutputToolStripMenuItem";
            this.audioOutputToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.audioOutputToolStripMenuItem.Text = "Audio Output";
            // 
            // breakerToolStripMenuItem
            // 
            this.breakerToolStripMenuItem.Name = "breakerToolStripMenuItem";
            this.breakerToolStripMenuItem.Size = new System.Drawing.Size(113, 22);
            this.breakerToolStripMenuItem.Text = "breaker";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(175, 6);
            // 
            // videoCompressorsToolStripMenuItem
            // 
            this.videoCompressorsToolStripMenuItem.Name = "videoCompressorsToolStripMenuItem";
            this.videoCompressorsToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.videoCompressorsToolStripMenuItem.Text = "Video Compressors";
            // 
            // audioCompressorsToolStripMenuItem
            // 
            this.audioCompressorsToolStripMenuItem.Name = "audioCompressorsToolStripMenuItem";
            this.audioCompressorsToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            this.audioCompressorsToolStripMenuItem.Text = "Audio Compressors";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(175, 6);
            // 
            // scanDevicesToolStripMenuItem
            // 
            this.scanDevicesToolStripMenuItem.Name = "scanDevicesToolStripMenuItem";
            this.scanDevicesToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.scanDevicesToolStripMenuItem.Text = "Scan Devices";
            this.scanDevicesToolStripMenuItem.Click += new System.EventHandler(this.MnuScan_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.videoSourcesToolStripMenuItem,
            this.videoFrameSizeToolStripMenuItem,
            this.videoFrameRateToolStripMenuItem,
            this.videoCapabilitiesToolStripMenuItem,
            this.toolStripSeparator3,
            this.audioSourcesToolStripMenuItem,
            this.audioChannelsToolStripMenuItem,
            this.audioSamplingRateToolStripMenuItem,
            this.audioSampleSizeToolStripMenuItem,
            this.audioCapabilitiesToolStripMenuItem,
            this.toolStripSeparator4,
            this.tVTunerChannelToolStripMenuItem,
            this.tVTunerInputTypeToolStripMenuItem,
            this.toolStripSeparator5,
            this.propertyPagesToolStripMenuItem,
            this.toolStripSeparator6,
            this.previewToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.optionsToolStripMenuItem.Text = "&Options";
            // 
            // videoSourcesToolStripMenuItem
            // 
            this.videoSourcesToolStripMenuItem.Name = "videoSourcesToolStripMenuItem";
            this.videoSourcesToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.videoSourcesToolStripMenuItem.Text = "Video Sources";
            // 
            // videoFrameSizeToolStripMenuItem
            // 
            this.videoFrameSizeToolStripMenuItem.Name = "videoFrameSizeToolStripMenuItem";
            this.videoFrameSizeToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.videoFrameSizeToolStripMenuItem.Text = "Video Frame Size";
            // 
            // videoFrameRateToolStripMenuItem
            // 
            this.videoFrameRateToolStripMenuItem.Name = "videoFrameRateToolStripMenuItem";
            this.videoFrameRateToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.videoFrameRateToolStripMenuItem.Text = "Video Frame Rate";
            // 
            // videoCapabilitiesToolStripMenuItem
            // 
            this.videoCapabilitiesToolStripMenuItem.Name = "videoCapabilitiesToolStripMenuItem";
            this.videoCapabilitiesToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.videoCapabilitiesToolStripMenuItem.Text = "Video Capabilities...";
            this.videoCapabilitiesToolStripMenuItem.Click += new System.EventHandler(this.VideoCapabilitiesToolStripMenuItem_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(182, 6);
            // 
            // audioSourcesToolStripMenuItem
            // 
            this.audioSourcesToolStripMenuItem.Name = "audioSourcesToolStripMenuItem";
            this.audioSourcesToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.audioSourcesToolStripMenuItem.Text = "Audio Sources";
            // 
            // audioChannelsToolStripMenuItem
            // 
            this.audioChannelsToolStripMenuItem.Name = "audioChannelsToolStripMenuItem";
            this.audioChannelsToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.audioChannelsToolStripMenuItem.Text = "Audio Channels";
            // 
            // audioSamplingRateToolStripMenuItem
            // 
            this.audioSamplingRateToolStripMenuItem.Name = "audioSamplingRateToolStripMenuItem";
            this.audioSamplingRateToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.audioSamplingRateToolStripMenuItem.Text = "Audio Sampling Rate";
            // 
            // audioSampleSizeToolStripMenuItem
            // 
            this.audioSampleSizeToolStripMenuItem.Name = "audioSampleSizeToolStripMenuItem";
            this.audioSampleSizeToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.audioSampleSizeToolStripMenuItem.Text = "Audio Sample Size";
            // 
            // audioCapabilitiesToolStripMenuItem
            // 
            this.audioCapabilitiesToolStripMenuItem.Name = "audioCapabilitiesToolStripMenuItem";
            this.audioCapabilitiesToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.audioCapabilitiesToolStripMenuItem.Text = "Audio Capabilities...";
            this.audioCapabilitiesToolStripMenuItem.Click += new System.EventHandler(this.AudioCapabilitiesToolStripMenuItem_Click);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(182, 6);
            // 
            // tVTunerChannelToolStripMenuItem
            // 
            this.tVTunerChannelToolStripMenuItem.Name = "tVTunerChannelToolStripMenuItem";
            this.tVTunerChannelToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.tVTunerChannelToolStripMenuItem.Text = "TV Tuner Channel";
            // 
            // tVTunerInputTypeToolStripMenuItem
            // 
            this.tVTunerInputTypeToolStripMenuItem.Name = "tVTunerInputTypeToolStripMenuItem";
            this.tVTunerInputTypeToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.tVTunerInputTypeToolStripMenuItem.Text = "TV Tuner Input Type";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(182, 6);
            // 
            // propertyPagesToolStripMenuItem
            // 
            this.propertyPagesToolStripMenuItem.Name = "propertyPagesToolStripMenuItem";
            this.propertyPagesToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.propertyPagesToolStripMenuItem.Text = "Property Pages";
            this.propertyPagesToolStripMenuItem.Click += new System.EventHandler(this.PropertyPagesToolStripMenuItem_Click);
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(182, 6);
            // 
            // previewToolStripMenuItem
            // 
            this.previewToolStripMenuItem.Name = "previewToolStripMenuItem";
            this.previewToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
            this.previewToolStripMenuItem.Text = "&Preview";
            this.previewToolStripMenuItem.Click += new System.EventHandler(this.PreviewToolStripMenuItem_Click);
            // 
            // toolStripSeparator7
            // 
            this.toolStripSeparator7.Name = "toolStripSeparator7";
            this.toolStripSeparator7.Size = new System.Drawing.Size(177, 6);
            // 
            // FrmCaptureGame
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(496, 385);
            this.Controls.Add(this.panelVideo);
            this.Controls.Add(this.msMainMenu);
            this.MainMenuStrip = this.msMainMenu;
            this.Menu = this.mainMenu;
            this.Name = "FrmCaptureGame";
            this.Text = "FrmCaptureTest";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CaptureTest_FormClosed);
            this.Load += new System.EventHandler(this.CaptureTest_Load);
            this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.FrmCaptureTest_KeyPress);
            this.msMainMenu.ResumeLayout(false);
            this.msMainMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            Application.Run(new FrmCaptureGame());
        }
        private void btnExit_Click(object sender, System.EventArgs e)
        {
            if (capture != null)
                capture.Stop();
            Application.Exit();
        }
        private void updateMenu()
        {
            ToolStripMenuItem tsmi =
                new ToolStripMenuItem("(None)", null, new EventHandler(mnuVideoDevices_Click));
            Filter f;
            Source s;
            Source current;
            PropertyPage p;
            Control oldPreviewWindow = null;

            // Disable preview to avoid additional flashes (optional)
            if (capture != null)
            {
                oldPreviewWindow = capture.PreviewWindow;
                capture.PreviewWindow = null;
            }

            // Load video devices
            Filter videoDevice = null;
            if (capture != null)
                videoDevice = capture.VideoDevice;
            videoDevicesToolStripMenuItem.DropDownItems.Clear();

            tsmi.Checked = (videoDevice == null);
            videoDevicesToolStripMenuItem.DropDownItems.Add(tsmi);
            for (int c = 0; c < filters.VideoInputDevices.Count; c++)
            {
                f = filters.VideoInputDevices[c];
                tsmi = new ToolStripMenuItem(f.Name, null, new EventHandler(mnuVideoDevices_Click));
                tsmi.Checked = (videoDevice == f);
                videoDevicesToolStripMenuItem.DropDownItems.Add(tsmi);
            }
            videoDevicesToolStripMenuItem.Enabled = (filters.VideoInputDevices.Count > 0);

            // Load audio devices
            Filter audioDevice = null;
            if (capture != null)
                audioDevice = capture.AudioDevice;
            audioDevicesToolStripMenuItem.DropDownItems.Clear();
            tsmi = new ToolStripMenuItem("(None)", null, new EventHandler(mnuAudioDevices_Click));
            tsmi.Checked = (audioDevice == null);
            audioDevicesToolStripMenuItem.DropDownItems.Add(tsmi);
            for (int c = 0; c < filters.AudioInputDevices.Count; c++)
            {
                f = filters.AudioInputDevices[c];
                tsmi = new ToolStripMenuItem(f.Name, null, new EventHandler(mnuAudioDevices_Click));
                tsmi.Checked = (audioDevice == f);
                audioDevicesToolStripMenuItem.DropDownItems.Add(tsmi);
            }
            audioDevicesToolStripMenuItem.Enabled = (filters.AudioInputDevices.Count > 0);

            // Load Output device list
            try
            {
                devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
                string CheckedDeviceName = ""; // String to hold the checked device name
                foreach(ToolStripMenuItem i in audioOutputToolStripMenuItem.DropDownItems)
                {
                    //If the item is checked, store the checked item name to the string
                    if (i.Checked)
                    { CheckedDeviceName = i.Text; }
                }
                /*for (int i = 0; i < audioOutputToolStripMenuItem.DropDownItems.Count; i++)
                {
                    tsmi = (ToolStripMenuItem)audioOutputToolStripMenuItem.DropDownItems[i];
                    //If the item is checked, store the checked item to the string
                    if (tsmi.Checked)
                    { CheckedDeviceName = devices[i].Name; }
                }/**/
                // Set the first device as Output on the first run
                if (CheckedDeviceName.Equals(""))
                { CheckedDeviceName = devices[0].Name; }

                audioOutputToolStripMenuItem.DropDownItems.Clear();
                for (int c = 0; c < devices.Length; c++)
                {   /// TODO: work out which devices are usable and process accodingly
                    /// Some devices show as DirectShow etc

                    MenuItemAdd(devices[c].Name, Resources.IconAudio.ToBitmap(), new EventHandler(MnuAudioOutput_Click),
                        (devices[c].Name.Equals(CheckedDeviceName)), audioOutputToolStripMenuItem);
                }
                audioOutputToolStripMenuItem.Enabled = true;
            }
            catch { audioOutputToolStripMenuItem.Enabled = false; }
            // Load video compressors
            try
            {
                videoCompressorsToolStripMenuItem.DropDownItems.Clear();
                tsmi = new ToolStripMenuItem("(None)", null, new EventHandler(mnuVideoCompressors_Click));
                tsmi.Checked = (capture.VideoCompressor == null);
                videoCompressorsToolStripMenuItem.DropDownItems.Add(tsmi);
                for (int c = 0; c < filters.VideoCompressors.Count; c++)
                {
                    f = filters.VideoCompressors[c];
                    tsmi = new ToolStripMenuItem(f.Name, null, new EventHandler(mnuVideoCompressors_Click));
                    tsmi.Checked = (capture.VideoCompressor == f);
                    videoCompressorsToolStripMenuItem.DropDownItems.Add(tsmi);
                }

                videoCompressorsToolStripMenuItem.Enabled = ((capture.VideoDevice != null) && (filters.VideoCompressors.Count > 0));
            }
            catch { videoCompressorsToolStripMenuItem.Enabled = false; }

            // Load audio compressors
            try
            {
                audioCompressorsToolStripMenuItem.DropDownItems.Clear();

                tsmi = new ToolStripMenuItem("(None)", null, new EventHandler(mnuAudioCompressors_Click));
                tsmi.Checked = (capture.AudioCompressor == null);
                audioCompressorsToolStripMenuItem.DropDownItems.Add(tsmi);
                for (int c = 0; c < filters.AudioCompressors.Count; c++)
                {
                    f = filters.AudioCompressors[c];
                    tsmi = new ToolStripMenuItem(f.Name, null, new EventHandler(mnuAudioCompressors_Click));
                    tsmi.Checked = (capture.AudioCompressor == f);
                    audioCompressorsToolStripMenuItem.DropDownItems.Add(tsmi);

                }
                audioCompressorsToolStripMenuItem.Enabled = ((capture.AudioDevice != null) && (filters.AudioCompressors.Count > 0));
            }
            catch { audioCompressorsToolStripMenuItem.Enabled = false; }

            // Load video sources
            try
            {
                videoSourcesToolStripMenuItem.DropDownItems.Clear();
                current = capture.VideoSource;
                for (int c = 0; c < capture.VideoSources.Count; c++)
                {
                    s = capture.VideoSources[c];
                    tsmi = new ToolStripMenuItem(s.Name, null, new EventHandler(mnuVideoSources_Click));
                    tsmi.Checked = (current == s);
                    videoSourcesToolStripMenuItem.DropDownItems.Add(tsmi);
                }
                videoSourcesToolStripMenuItem.Enabled = (capture.VideoSources.Count > 0);
            }
            catch { videoSourcesToolStripMenuItem.Enabled = false; }

            // Load audio sources
            try
            {
                audioSourcesToolStripMenuItem.DropDownItems.Clear();
                current = capture.AudioSource;
                for (int c = 0; c < capture.AudioSources.Count; c++)
                {
                    s = capture.AudioSources[c];
                    MenuItemAdd(s.Name, new EventHandler(mnuAudioSources_Click), (current == s), audioSourcesToolStripMenuItem);
                }
                audioSourcesToolStripMenuItem.Enabled = (capture.AudioSources.Count > 0);
            }
            catch { audioSourcesToolStripMenuItem.Enabled = false; }

            // Load frame rates
            try
            {
                videoFrameRateToolStripMenuItem.DropDownItems.Clear();

                int frameRate = (int)(capture.FrameRate * 1000);
                MenuItemAdd("15 fps", new EventHandler(mnuFrameRates_Click), (frameRate == 15000), videoFrameRateToolStripMenuItem);
                MenuItemAdd("24 fps (Film)", new EventHandler(mnuFrameRates_Click),
                    (frameRate == 24000), videoFrameRateToolStripMenuItem);
                tsmi = new ToolStripMenuItem("25 fps (PAL)", null, new EventHandler(mnuFrameRates_Click));
                tsmi.Checked = (frameRate == 25000);
                videoFrameRateToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("29.997 fps (NTSC)", null, new EventHandler(mnuFrameRates_Click));
                tsmi.Checked = (frameRate == 29997);
                videoFrameRateToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("30 fps (~NTSC)", null, new EventHandler(mnuFrameRates_Click));
                tsmi.Checked = (frameRate == 30000);
                videoFrameRateToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("59.994 fps (2xNTSC)", null, new EventHandler(mnuFrameRates_Click));
                tsmi.Checked = (frameRate == 59994);
                videoFrameRateToolStripMenuItem.DropDownItems.Add(tsmi);
                MenuItemAdd("60 FPS", new EventHandler(mnuFrameRates_Click), (frameRate == 60000), videoFrameRateToolStripMenuItem);
                videoFrameRateToolStripMenuItem.Enabled = true;
            }
            catch { videoFrameRateToolStripMenuItem.Enabled = false; }

            // Load frame sizes
            try
            {
                videoFrameSizeToolStripMenuItem.DropDownItems.Clear();
                Size frameSize = capture.FrameSize;
                tsmi = new ToolStripMenuItem("160 x 120", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(160, 120));
                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("320 x 240", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(320, 240));
                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("640 x 480", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(640, 480));
                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("1024 x 768", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(1024, 768));
                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("1280 x 720", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(1280, 720));
                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("1920 x 1080", null, new EventHandler(mnuFrameSizes_Click));
                tsmi.Checked = (frameSize == new Size(1920, 1080));
                tsmi.Enabled = (frameSize == new Size(1920, 1080)) ? true : false;

                videoFrameSizeToolStripMenuItem.DropDownItems.Add(tsmi);

                videoFrameSizeToolStripMenuItem.Enabled = true;
            }
            catch { videoFrameSizeToolStripMenuItem.Enabled = false; }

            // Load audio channels
            try
            {
                audioChannelsToolStripMenuItem.DropDownItems.Clear();
                short audioChannels = capture.AudioChannels;
                tsmi = new ToolStripMenuItem("Mono", null, new EventHandler(mnuAudioChannels_Click));
                tsmi.Checked = (audioChannels == 1);
                audioChannelsToolStripMenuItem.DropDownItems.Add(tsmi);
                tsmi = new ToolStripMenuItem("Stereo", null, new EventHandler(mnuAudioChannels_Click));
                tsmi.Checked = (audioChannels == 2);
                audioChannelsToolStripMenuItem.DropDownItems.Add(tsmi);
                audioChannelsToolStripMenuItem.Enabled = true;

            }
            catch { audioChannelsToolStripMenuItem.Enabled = false; }

            // Load audio sampling rate
            try
            {
                audioSamplingRateToolStripMenuItem.DropDownItems.Clear();
                int samplingRate = capture.AudioSamplingRate;
                MenuItemAdd("8 kHz", new EventHandler(mnuAudioSamplingRate_Click),
                    (samplingRate == 8000), audioSamplingRateToolStripMenuItem);
                MenuItemAdd("11.025 kHz", new EventHandler(mnuAudioSamplingRate_Click),
                    (capture.AudioSamplingRate == 11025), audioSamplingRateToolStripMenuItem);
                MenuItemAdd("22.05 kHz", new EventHandler(mnuAudioSamplingRate_Click),
                    (capture.AudioSamplingRate == 22050), audioSamplingRateToolStripMenuItem);
                MenuItemAdd("44.1 kHz", null, new EventHandler(mnuAudioSamplingRate_Click),
                (capture.AudioSamplingRate == 44100), audioSamplingRateToolStripMenuItem);
                audioSamplingRateToolStripMenuItem.Enabled = true;
            }
            catch { audioSamplingRateToolStripMenuItem.Enabled = false; }

            // Load audio sample sizes
            try
            {
                audioSampleSizeToolStripMenuItem.DropDownItems.Clear();
                short sampleSize = capture.AudioSampleSize;
                MenuItemAdd("8 bit", new EventHandler(mnuAudioSampleSizes_Click), (sampleSize == 8), audioSampleSizeToolStripMenuItem);
                MenuItemAdd("16 bit", new EventHandler(mnuAudioSampleSizes_Click), (sampleSize == 16), audioSampleSizeToolStripMenuItem);
                //MenuItemAdd("32 bit", new EventHandler(mnuAudioSampleSizes_Click), (sampleSize == 32), audioSampleSizeToolStripMenuItem);
                audioSampleSizeToolStripMenuItem.Enabled = true;
            }
            catch { audioSampleSizeToolStripMenuItem.Enabled = false; }

            // Load property pages
            try
            {
                propertyPagesToolStripMenuItem.DropDownItems.Clear();
                for (int c = 0; c < capture.PropertyPages.Count; c++)
                {
                    p = capture.PropertyPages[c];
                    MenuItemAdd(p.Name + "...", new EventHandler(mnuPropertyPages_Click), false, propertyPagesToolStripMenuItem);

                }
                propertyPagesToolStripMenuItem.Enabled = (capture.PropertyPages.Count > 0);
            }
            catch { propertyPagesToolStripMenuItem.Enabled = false; }

            // Load TV Tuner channels
            try
            {
                tVTunerChannelToolStripMenuItem.DropDownItems.Clear();
                int channel = capture.Tuner.Channel;
                for (int c = 1; c <= 25; c++)
                {
                    MenuItemAdd(c.ToString(), new EventHandler(mnuChannel_Click), (channel == c), tVTunerChannelToolStripMenuItem);
                }
                tVTunerChannelToolStripMenuItem.Enabled = true;
            }
            catch { tVTunerChannelToolStripMenuItem.Enabled = false; }

            // Load TV Tuner input types
            try
            {
                tVTunerInputTypeToolStripMenuItem.DropDownItems.Clear();
                MenuItemAdd(TunerInputType.Cable.ToString(), new EventHandler(mnuInputType_Click),
                    (capture.Tuner.InputType == TunerInputType.Cable), tVTunerInputTypeToolStripMenuItem);
                MenuItemAdd(TunerInputType.Antenna.ToString(), new EventHandler(mnuInputType_Click),
                    (capture.Tuner.InputType == TunerInputType.Antenna), tVTunerInputTypeToolStripMenuItem);


                tVTunerInputTypeToolStripMenuItem.Enabled = true;
            }
            catch { tVTunerInputTypeToolStripMenuItem.Enabled = false; }

            // Enable/disable caps
            videoCapabilitiesToolStripMenuItem.Enabled = ((capture != null) && (capture.VideoCaps != null));
            audioCapabilitiesToolStripMenuItem.Enabled = ((capture != null) && (capture.AudioCaps != null));

            // Check Preview menu option
            previewToolStripMenuItem.Checked = (oldPreviewWindow != null);
            previewToolStripMenuItem.Enabled = (capture != null);

            // Reenable preview if it was enabled before
            if (capture != null)
                capture.PreviewWindow = oldPreviewWindow;
        }

        private void mnuVideoDevices_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Get current devices and dispose of capture object
                // because the video and audio device can only be changed
                // by creating a new Capture object.
                Filter videoDevice = null;
                Filter audioDevice = null;
                if (capture != null)
                {
                    videoDevice = capture.VideoDevice;
                    audioDevice = capture.AudioDevice;
                    capture.Dispose();
                    capture = null;
                }

                // Get new video device
                ToolStripMenuItem m = sender as ToolStripMenuItem;

                videoDevice = (m.Owner.Items.Count > 0 ? filters.VideoInputDevices[m.Owner.Items.IndexOf(m) - 1] : null);
                //MenuItem m = sender as MenuItem;
                //videoDevice = (  m.Index>0 ? filters.VideoInputDevices[m.Index-1] : null );

                // Create capture object
                if ((videoDevice != null) || (audioDevice != null))
                {
                    capture = new Capture(videoDevice, audioDevice);
                    capture.CaptureComplete += new EventHandler(OnCaptureComplete);
                }

                // Update the menu
                updateMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Video device not supported.\n\n" + ex.Message + "\n\n" + ex.ToString());
            }
        }

        private void mnuAudioDevices_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Get current devices and dispose of capture object
                // because the video and audio device can only be changed
                // by creating a new Capture object.
                Filter videoDevice = null;
                Filter audioDevice = null;
                if (capture != null)
                {
                    videoDevice = capture.VideoDevice;
                    audioDevice = capture.AudioDevice;
                    capture.Dispose();
                    capture = null;
                }

                // Get new audio device
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                audioDevice = (m.Owner.Items.Count > 0 ? filters.AudioInputDevices[m.Owner.Items.IndexOf(m) - 1] : null);

                // Create capture object
                if ((videoDevice != null) || (audioDevice != null))
                {
                    capture = new Capture(videoDevice, audioDevice);
                    capture.CaptureComplete += new EventHandler(OnCaptureComplete);
                }

                // Update the menu
                updateMenu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Audio device not supported.\n\n" + ex.Message + "\n\n" + ex.ToString());
            }
        }
        private void MnuAudioOutput_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Get new audio output
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                m.Checked = true;

                // Update the menu
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Output Error:\n\n" + ex.Message + "\n\n" + ex.ToString(), "Output ERROR"); }
        }
        private void mnuVideoCompressors_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Change the video compressor
                // We subtract 1 from m.Index beacuse the first item is (None)
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.VideoCompressor = (m.Owner.Items.Count > 0 ? filters.VideoCompressors[m.Owner.Items.IndexOf(m) - 1] : null);

                //MenuItem m = sender as MenuItem;
                //capture.VideoCompressor = ( m.Index>0 ? filters.VideoCompressors[m.Index-1] : null );
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Video compressor not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuAudioCompressors_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Change the audio compressor
                // We subtract 1 from m.Index beacuse the first item is (None)
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.AudioCompressor = (m.Owner.Items.IndexOf(m) > 0 ? filters.AudioCompressors[m.Owner.Items.IndexOf(m) - 1] : null);
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Audio compressor not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuVideoSources_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Choose the video source
                // If the device only has one source, this menu item will be disabled
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.VideoSource = capture.VideoSources[m.Owner.Items.IndexOf(m)];
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Unable to set video source.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuAudioSources_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Choose the audio source
                // If the device only has one source, this menu item will be disabled
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.AudioSource = capture.AudioSources[m.Owner.Items.IndexOf(m)];
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Unable to set audio source.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }


        private void mnuExit_Click(object sender, System.EventArgs e)
        {
            Exit_App();
        }

        private void mnuFrameSizes_Click(object sender, System.EventArgs e)
        {
            try
            {
                // Disable preview to avoid additional flashes (optional)
                bool preview = (capture.PreviewWindow != null);
                capture.PreviewWindow = null;

                // Update the frame size
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                string[] s = m.Text.Split('x');
                Size size = new Size(int.Parse(s[0]), int.Parse(s[1]));
                capture.FrameSize = size;

                // Update the menu
                updateMenu();

                // Restore previous preview setting
                capture.PreviewWindow = (preview ? panelVideo : null);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Frame size not supported.\n\n" + ex.Message + "\n\n" + ex.ToString(),
                    "Frame size not supported", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void mnuFrameRates_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                string[] s = m.Text.Split(' ');
                capture.FrameRate = double.Parse(s[0]);
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Frame rate not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }


        private void mnuAudioChannels_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.AudioChannels = (short)Math.Pow(2, m.Owner.Items.IndexOf(m) + 1);
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Number of audio channels not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuAudioSamplingRate_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                string[] s = m.Text.Split(' ');
                int samplingRate = (int)(double.Parse(s[0]) * 1000);
                capture.AudioSamplingRate = samplingRate;
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Audio sampling rate not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuAudioSampleSizes_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                string[] s = m.Text.Split(' ');
                short sampleSize = short.Parse(s[0]);
                capture.AudioSampleSize = sampleSize;
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Audio sample size not supported.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuPreview_Click(object sender, System.EventArgs e)
        {
            try
            {
                if (capture.PreviewWindow == null)
                {
                    capture.PreviewWindow = panelVideo;
                    previewToolStripMenuItem.Checked = true;
                }
                else
                {
                    capture.PreviewWindow = null;
                    previewToolStripMenuItem.Checked = false;
                }
            }
            catch (Exception ex)
            { MessageBox.Show("Unable to enable/disable preview.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuPropertyPages_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.PropertyPages[m.Owner.Items.IndexOf(m)].Show(this);
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Unable display property page.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuChannel_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.Tuner.Channel = m.Owner.Items.IndexOf(m) + 1;
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Unable change channel.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void mnuInputType_Click(object sender, System.EventArgs e)
        {
            try
            {
                ToolStripMenuItem m = sender as ToolStripMenuItem;
                capture.Tuner.InputType = (TunerInputType)m.Owner.Items.IndexOf(m);
                updateMenu();
            }
            catch (Exception ex)
            { MessageBox.Show("Unable change tuner input type.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void OnCaptureComplete(object sender, EventArgs e)
        {
            // Demonstrate the Capture.CaptureComplete event.
            Debug.WriteLine("Capture complete.");
        }

        private void PanelVideo_Paint(object sender, PaintEventArgs e)
        {
        }

        private void CaptureTest_Load(object sender, EventArgs e)
        {
        }

        private void CaptureTest_FormClosed(object sender, FormClosedEventArgs e)
        {
            Exit_App();
        }

        private void Exit_App()
        {
            string selectedVD = "", selectedARO = "", selectedAID = "";
            
            // get sound output devices
            DsDevice[] devices;
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
            foreach (ToolStripMenuItem i in audioOutputToolStripMenuItem.DropDownItems)
            { // Search the output device menu for the selected device and get it's index number
                if (i.Checked)
                {
                    foreach (DsDevice dsd in devices)
                    {
                        if (dsd.Name.Equals(i.Text))
                        {
                            selectedARO = dsd.Name;
                            break;
                        }
                    }
                }
            }

            // Mic input
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioInputDevice);
            foreach (ToolStripMenuItem i in audioDevicesToolStripMenuItem.DropDownItems)
            {  // find the desired input device to be used from the AudioDevices menu
                if (i.Checked)
                { selectedAID = i.Text; }
            }
            // Save user settings
            Settings.Default.LastAudioRenderer = selectedARO;
            Settings.Default.LastAudioIn = selectedAID;
            Settings.Default.LastVideoDevice = selectedVD;
            Settings.Default.Save();
            // Kill everything running
            m_objMediaControl.Stop();
            Marshal.ReleaseComObject(m_objFilterGraph);
            m_objMediaControl = null;
            m_objBasicAudio = null;
            if (capture != null)
            { capture.Stop(); }
            components.Dispose();
            Application.Exit();
        }

        private void PanelVideo_MouseClick(object sender, MouseEventArgs e)
        {
         
        }

        private void Hide_panel()
        {
            //hide
            if (msMainMenu.Visible)
            { msMainMenu.Hide(); }
            else
            { msMainMenu.Show(); }
            FrmCaptureGame.ActiveForm.Activate();

            // Change to borderless
            if (FrmCaptureGame.ActiveForm.FormBorderStyle == FormBorderStyle.None)
            { FrmCaptureGame.ActiveForm.FormBorderStyle = FormBorderStyle.FixedSingle; }
            else
            {
                FrmCaptureGame.ActiveForm.FormBorderStyle = FormBorderStyle.None;
                FrmCaptureGame.ActiveForm.Height = 1080;
                FrmCaptureGame.ActiveForm.Width = 1920;
            }
        }

        private void MnuScan_Click(object sender, EventArgs e)
        {
            updateMenu();
        }

        /// <summary>
        /// Adds menu items to the specified DropDownItems target 
        /// </summary>
        /// <param name="itemText">Text for the menu item.<./param>
        /// <param name="ehHandle">EventHandler the item will be assigned.</param>
        /// <param name="makeChecked">Make the menu item checked</param>
        /// <param name="dropDownTarget">ToolStripMenuItem to put the menu item under</param>
        private void MenuItemAdd(string itemText, EventHandler ehHandle, bool makeChecked, ToolStripMenuItem dropDownTarget)
        {
            ToolStripMenuItem tsmi = new ToolStripMenuItem(itemText, null, ehHandle){ Checked = makeChecked };
            dropDownTarget.DropDownItems.Add(tsmi);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="itemText">Text for the menu item.</param>
        /// <param name="mnuImage">The image that appears next to the item, like an icon</param>
        /// <param name="ehHandle">EventHandler the item will be assigned.</param>
        /// <param name="makeChecked">Make the menu item checked</param>
        /// <param name="dropDownTarget">ToolStripMenuItem to put the menu item under</param>
        private void MenuItemAdd(string itemText, Image mnuImage, EventHandler ehHandle, bool makeChecked, ToolStripMenuItem dropDownTarget)
        {
            ToolStripMenuItem tsmi = new ToolStripMenuItem(itemText, mnuImage, ehHandle){ Checked = makeChecked };
            dropDownTarget.DropDownItems.Add(tsmi);
        }
        private void PreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (capture.PreviewWindow == null)
                {
                    capture.PreviewWindow = panelVideo;
                    previewToolStripMenuItem.Checked = true;

                    ///TODO: Sync issue?

                    string source = null ;
                    int readIn = 0, readOut = 0;

                    // get sound output devices
                    DsDevice[] devices;
                    devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioRendererCategory);
                    int count = 0;
                    foreach (ToolStripMenuItem i in audioOutputToolStripMenuItem.DropDownItems)
                    { // Search the output device menu for the selected device and get it's index number
                        if (i.Checked)
                        {
                            foreach(DsDevice dsd in devices)
                            {
                                if (dsd.Name.Equals(i.Text))
                                {
                                    readOut = count;
                                    break;
                                }
                            }
                        }
                        count++;
                    }
                   
                    // Mic input
                    devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioInputDevice);
                    count = 0;
                    
                    foreach (ToolStripMenuItem i in audioDevicesToolStripMenuItem.DropDownItems)
                    {  // find the desired input device to be used from the AudioDevices menu
                        if (i.Checked)
                        { source = i.Text; }
                    }
                    //Find the checked device in the 'devices' list and get the index integer of it
                    foreach (DsDevice dsd in devices)
                    {
                        if (dsd.Name.Equals(source))
                        {
                            readIn = count;
                            break;
                        }
                        count++;
                    }
                    // Connect input to output then run
                    Connect_up(readIn,readOut);
                    m_objMediaControl.Run();
                }
                else
                {
                    // Kill everything
                    capture.PreviewWindow = null;
                    previewToolStripMenuItem.Checked = false;
                    // Stop graph
                    m_objMediaControl.Stop();
                    m_objFilterGraph.Abort();
                }
            }
            catch (Exception ex)
            { MessageBox.Show("Unable to enable/disable preview.\n\n" + ex.Message + "\n\n" + ex.ToString()); }
        }

        private void FrmCaptureTest_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch (e.KeyChar)
            {
                case 'h':
                    Hide_panel();
                    break;
                case 'p':
                    PreviewToolStripMenuItem_Click(this, null);
                    break;
            }
                
            
        }

        private void StartPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            capture.Start();
        }

        private void StopPreviewToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void VideoCapabilitiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string s;
                s = String.Format(
                    "Video Device Capabilities\n" +
                    "--------------------------------\n\n" +
                    "Input Size:\t\t{0} x {1}\n" +
                    "\n" +
                    "Min Frame Size:\t\t{2} x {3}\n" +
                    "Max Frame Size:\t\t{4} x {5}\n" +
                    "Frame Size Granularity X:\t{6}\n" +
                    "Frame Size Granularity Y:\t{7}\n" +
                    "\n" +
                    "Min Frame Rate:\t\t{8:0.000} fps\n" +
                    "Max Frame Rate:\t\t{9:0.000} fps\n",
                    capture.VideoCaps.InputSize.Width, capture.VideoCaps.InputSize.Height,
                    capture.VideoCaps.MinFrameSize.Width, capture.VideoCaps.MinFrameSize.Height,
                    capture.VideoCaps.MaxFrameSize.Width, capture.VideoCaps.MaxFrameSize.Height,
                    capture.VideoCaps.FrameSizeGranularityX,
                    capture.VideoCaps.FrameSizeGranularityY,
                    capture.VideoCaps.MinFrameRate,
                    capture.VideoCaps.MaxFrameRate);
                MessageBox.Show(s, "Video Capibilities");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable display video capabilities. Please submit a bug report.\n\n" +
                    ex.Message + "\n\n" + ex.ToString());
            }
        }

        private void AudioCapabilitiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string s;
                s = String.Format(
                    "Audio Device Capabilities\n" +
                    "--------------------------------\n\n" +
                    "Min Channels:\t\t{0}\n" +
                    "Max Channels:\t\t{1}\n" +
                    "Channels Granularity:\t{2}\n" +
                    "\n" +
                    "Min Sample Size:\t\t{3}\n" +
                    "Max Sample Size:\t\t{4}\n" +
                    "Sample Size Granularity:\t{5}\n" +
                    "\n" +
                    "Min Sampling Rate:\t\t{6}\n" +
                    "Max Sampling Rate:\t\t{7}\n" +
                    "Sampling Rate Granularity:\t{8}\n",
                    capture.AudioCaps.MinimumChannels,
                    capture.AudioCaps.MaximumChannels,
                    capture.AudioCaps.ChannelsGranularity,
                    capture.AudioCaps.MinimumSampleSize,
                    capture.AudioCaps.MaximumSampleSize,
                    capture.AudioCaps.SampleSizeGranularity,
                    capture.AudioCaps.MinimumSamplingRate,
                    capture.AudioCaps.MaximumSamplingRate,
                    capture.AudioCaps.SamplingRateGranularity);
                MessageBox.Show(s, "Audio Capibilities");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable display audio capabilities. Please submit a bug report.\n\n" +
                    ex.Message + "\n\n" + ex.ToString(), "Audio Error");
            }
        }

        private void FileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Exit_App();
        }

        private void PropertyPagesToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ScanDevicesToolStripMenuItem_Click_1(object sender, EventArgs e)
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

            // Mic Input
            devices = DsDevice.GetDevicesOfCat(FilterCategory.AudioInputDevice);
            device = devices[dInput];
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
                    newI.FindPin("Capture", out pIn);//Get hold of the pin as seen in GraphEdit}
                }
                IBaseFilter ibfO = (IBaseFilter)source;
                ibfO.EnumPins(out pOutputPin);//Enumerate the pin

                ibfO.FindPin("Audio Input pin (rendered)", out pOut);
                try
                {
                    pIn.Connect(pOut, null);  //Connect the input pin to output pin
                }
                catch (Exception ex)
                { MessageBox.Show(ex.Message, "Pin Out Failure"); }
            }
            catch (Exception ex)
            { MessageBox.Show(ex.Message, "Connection Failure"); }

            m_objBasicAudio = m_objFilterGraph as IBasicAudio;
            m_objMediaControl = m_objFilterGraph as IMediaControl;

        }

        private void GameNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string newName = "";
            ShowInputDialog(ref newName);
            Settings.Default.GameName = newName;
            this.Text = newName;
        }
        private static DialogResult ShowInputDialog(ref string input)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Name";

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            DialogResult result = inputBox.ShowDialog();
            input = textBox.Text;
            return result;
        }
    }
}
