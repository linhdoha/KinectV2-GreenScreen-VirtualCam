using System;
using System.Diagnostics;
using Microsoft.Kinect;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Threading.Tasks;
using System.ComponentModel;

namespace KinectCam
{
	public static class KinectHelper
	{
		class KinectCamApplicationContext : ApplicationContext
		{
			private NotifyIcon TrayIcon;
			private ContextMenuStrip TrayIconContextMenu;
			private ToolStripMenuItem MirroredMenuItem;
			private ToolStripMenuItem DesktopMenuItem;
			private ToolStripMenuItem ZoomMenuItem;
			private ToolStripMenuItem TrackHeadMenuItem;
			public KinectCamApplicationContext()
			{
                System.Windows.Forms.Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
				InitializeComponent();
				TrayIcon.Visible = true;
				//TrayIcon.ShowBalloonTip(30000);
			}

			private void InitializeComponent()
			{
                TrayIcon = new NotifyIcon();

				TrayIcon.BalloonTipIcon = ToolTipIcon.Info;
				TrayIcon.BalloonTipText =
				  "For options use this tray icon.";
				TrayIcon.BalloonTipTitle = "KinectCamV2";
				TrayIcon.Text = "KinectCam";

				TrayIcon.Icon = IconExtractor.Extract(117, false);

				TrayIcon.DoubleClick += TrayIcon_DoubleClick;

				TrayIconContextMenu = new ContextMenuStrip();
				MirroredMenuItem = new ToolStripMenuItem();
				DesktopMenuItem = new ToolStripMenuItem();
				ZoomMenuItem = new ToolStripMenuItem();
				TrackHeadMenuItem = new ToolStripMenuItem();
				TrayIconContextMenu.SuspendLayout();

				// 
				// TrayIconContextMenu
				// 
				this.TrayIconContextMenu.Items.AddRange(new ToolStripItem[] {
				this.MirroredMenuItem,
				this.DesktopMenuItem,
				this.ZoomMenuItem,
				this.TrackHeadMenuItem
				});
				this.TrayIconContextMenu.Name = "TrayIconContextMenu";
				this.TrayIconContextMenu.Size = new Size(153, 70);
				// 
				// MirroredMenuItem
				// 
				this.MirroredMenuItem.Name = "Mirrored";
				this.MirroredMenuItem.Size = new Size(152, 22);
				this.MirroredMenuItem.Text = "Mirrored";
				this.MirroredMenuItem.Click += new EventHandler(this.MirroredMenuItem_Click);

				// 
				// DesktopMenuItem
				// 
				this.DesktopMenuItem.Name = "Desktop";
				this.DesktopMenuItem.Size = new Size(152, 22);
				this.DesktopMenuItem.Text = "Desktop";
				this.DesktopMenuItem.Click += new EventHandler(this.DesktopMenuItem_Click);

				// 
				// ZoomMenuItem
				// 
				this.ZoomMenuItem.Name = "Zoom";
				this.ZoomMenuItem.Size = new Size(152, 22);
				this.ZoomMenuItem.Text = "Zoom";
				this.ZoomMenuItem.Click += new EventHandler(this.ZoomMenuItem_Click);

				// 
				// ZoomMenuItem
				//
				this.TrackHeadMenuItem.Name = "TrackHead";
				this.TrackHeadMenuItem.Size = new Size(152, 22);
				this.TrackHeadMenuItem.Text = "TrackHead";
				this.TrackHeadMenuItem.Click += new EventHandler(this.TrackHeadMenuItem_Click);

				TrayIconContextMenu.ResumeLayout(false);
				TrayIcon.ContextMenuStrip = TrayIconContextMenu;
			}

			private void OnApplicationExit(object sender, EventArgs e)
			{
				TrayIcon.Visible = false;
			}

			private void TrayIcon_DoubleClick(object sender, EventArgs e)
			{
				TrayIcon.ShowBalloonTip(30000);
			}

			private void MirroredMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Mirrored = !KinectCamSettigns.Default.Mirrored;
			}

			private void DesktopMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Desktop = !KinectCamSettigns.Default.Desktop;
			}

			private void ZoomMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.Zoom = !KinectCamSettigns.Default.Zoom;
			}

			private void TrackHeadMenuItem_Click(object sender, EventArgs e)
			{
				KinectCamSettigns.Default.TrackHead = !KinectCamSettigns.Default.TrackHead;
			}
			public void Exit()
			{
				TrayIcon.Visible = false;
			}
		}

        /// <summary>
        /// Size of the RGB pixel in the bitmap
        /// </summary>
        private static readonly int bytesPerPixel = (PixelFormats.Bgr32.BitsPerPixel + 7) / 8;

        /// <summary>
        /// Coordinate mapper to map one type of point to another
        /// </summary>
        private static CoordinateMapper coordinateMapper = null;

        /// <summary>
        /// Reader for depth/color/body index frames
        /// </summary>
        private static MultiSourceFrameReader multiFrameSourceReader = null;

        /// <summary>
        /// Bitmap to display
        /// </summary>
        private static WriteableBitmap bitmap = null;

        /// <summary>
        /// The size in bytes of the bitmap back buffer
        /// </summary>
        private static uint bitmapBackBufferSize = 0;

        /// <summary>
        /// Intermediate storage for the color to depth mapping
        /// </summary>
        private static DepthSpacePoint[] colorMappedToDepthPoints = null;

        /// <summary>
        /// Current status text to display
        /// </summary>
        private static string statusText = null;



        static KinectCamApplicationContext context;
		static Thread contexThread;
		static Thread refreshThread;
		static KinectSensor Sensor;

        static void InitializeSensor()
		{
            var sensor = Sensor;
			if (sensor != null) return;

			try
			{
				sensor = KinectSensor.GetDefault();
				if (sensor == null) return;

				var reader = sensor.OpenMultiSourceFrameReader(FrameSourceTypes.Depth | FrameSourceTypes.Color | FrameSourceTypes.BodyIndex);
                reader.MultiSourceFrameArrived += Reader_MultiSourceFrameArrived;

                coordinateMapper = sensor.CoordinateMapper;

                FrameDescription depthFrameDescription = sensor.DepthFrameSource.FrameDescription;

                int depthWidth = depthFrameDescription.Width;
                int depthHeight = depthFrameDescription.Height;

                /*FrameDescription colorFrameDescription = sensor.ColorFrameSource.FrameDescription;

                int colorWidth = colorFrameDescription.Width;
                int colorHeight = colorFrameDescription.Height;

                colorMappedToDepthPoints = new DepthSpacePoint[colorWidth * colorHeight];

                bitmap = new WriteableBitmap(colorWidth, colorHeight, 96.0, 96.0, PixelFormats.Bgra32, null);

                // Calculate the WriteableBitmap back buffer size
                bitmapBackBufferSize = (uint)((bitmap.BackBufferStride * (bitmap.PixelHeight - 1)) + (bitmap.PixelWidth * bytesPerPixel));*/


                sensor.Open();

				Sensor = sensor;

				if (context == null)
				{
					contexThread = new Thread(() =>
					{
						context = new KinectCamApplicationContext();
						Application.Run(context);
					});
					refreshThread = new Thread(() =>
					{
						while (true)
						{
							Thread.Sleep(250);
							Application.DoEvents();
						}
					});
					contexThread.IsBackground = true;
					refreshThread.IsBackground = true;
					contexThread.SetApartmentState(ApartmentState.STA);
					refreshThread.SetApartmentState(ApartmentState.STA);
					contexThread.Start();
					refreshThread.Start();
				}
			}
			catch
			{
				Trace.WriteLine("Error of enable the Kinect sensor!");
			}

        }

		public delegate void InvokeDelegate();

		static void reader_FrameArrived(object sender, MultiSourceFrameArrivedEventArgs e)
		{
			var reference = e.FrameReference.AcquireFrame();
			using (var colorFrame = reference.ColorFrameReference.AcquireFrame())
			{
				if (colorFrame != null)
				{
					ColorFrameReady(colorFrame);
				}
			}
			using (var bodyFrame = reference.BodyFrameReference.AcquireFrame())
			{
				if (bodyFrame != null)
				{
					var _bodies = new Body[bodyFrame.BodyFrameSource.BodyCount];

					bodyFrame.GetAndRefreshBodyData(_bodies);
					_headFound = false;
					foreach (var body in _bodies)
					{
						if (body.IsTracked)
						{
							Joint head = body.Joints[JointType.Head];

							if (head.TrackingState == TrackingState.NotTracked)
								continue;
							_headFound = true;
							_headPosition = Sensor.CoordinateMapper.MapCameraPointToColorSpace(head.Position);

						}
					}
					UpdateZoomPosition();
				}
			}
		}
		static int SmoothTransitionByStep(int current, int needed, int step = 4)
		{

			if (step == 0)
			{
				return needed;
			}
			if (current > needed)
			{
				current -= step;
				if (current < needed)
					current = needed;
			}
			else if (current < needed)
			{
				current += step;
				if (current > needed)
					current = needed;
			}
			return current;
		}
		private static void UpdateZoomPosition()
		{
			// we should be at 30 fps in this place as Kinect Body are at 30 fps 
			int NeededZoomedWidthStart = 0;
			int NeededZoomedHeightStart;
			if (!KinectCamSettigns.Default.TrackHead || !_headFound)
			{
				NeededZoomedWidthStart = (SensorWidth - ZoomedWidth) / 2;
				NeededZoomedHeightStart = (SensorHeight - ZoomedHeight) / 2;
			}
			else
			{
				NeededZoomedWidthStart = (int)Math.Min(MaxZoomedWidthStart, Math.Max(0, _headPosition.X - ZoomedWidth / 2));

				NeededZoomedHeightStart = (int)Math.Min(MaxZoomedHeightStart, Math.Max(0, _headPosition.Y - ZoomedHeight / 2));
			}

			ZoomedWidthStart = SmoothTransitionByStep(ZoomedWidthStart, NeededZoomedWidthStart, 4);
			ZoomedHeightStart = SmoothTransitionByStep(ZoomedHeightStart, NeededZoomedHeightStart, 4);

			ZoomedWidthEnd = ZoomedWidthStart + ZoomedWidth;
			ZoomedHeightEnd = ZoomedHeightStart + ZoomedHeight;
			ZoomedPointerStart = ZoomedHeightStart * 1920 * 4 + ZoomedWidthStart * 4;
			ZoomedPointerEnd = ZoomedHeightEnd * 1920 * 4 + ZoomedWidthEnd * 4;

		}

		static unsafe void ColorFrameReady(ColorFrame frame)
		{
			if (frame.RawColorImageFormat == ColorImageFormat.Bgra)
			{
				frame.CopyRawFrameDataToArray(sensorColorFrameData);
			}
			else
			{
				frame.CopyConvertedFrameDataToArray(sensorColorFrameData, ColorImageFormat.Bgra);
			}
		}

		public static void DisposeSensor()
		{
			try
			{
				var sensor = Sensor;
				if (sensor != null && sensor.IsOpen)
				{
					sensor.Close();
					sensor = null;
					Sensor = null;
				}

				if (context != null)
				{
					context.Exit();
					context.Dispose();
					context = null;

					contexThread.Abort();
					refreshThread.Abort();
				}
			}
			catch
			{
				Trace.WriteLine("Error of disable the Kinect sensor!");
			}
		}

		static bool _headFound;
		static ColorSpacePoint _headPosition;

		public const int SensorWidth = 1920;
		public const int SensorHeight = 1080;
		public const int ZoomedWidth = 960;
		public const int ZoomedHeight = 540;

		public const int DefaultZoomedWidthStart = (SensorWidth - ZoomedWidth) / 2;
		public const int DefaultHeightStart = (SensorHeight - ZoomedHeight) / 2;

		public const int MaxZoomedWidthStart = SensorWidth - ZoomedWidth ;
		public const int MaxZoomedHeightStart = SensorHeight - (ZoomedHeight+1) ; // +1 to avoid overflow

		public static int ZoomedWidthStart = (SensorWidth - ZoomedWidth) / 2;
		public static int ZoomedHeightStart = (SensorHeight - ZoomedHeight) / 2;
		public static int ZoomedWidthEnd = ZoomedWidthStart + ZoomedWidth;
		public static int ZoomedHeightEnd = ZoomedHeightStart + ZoomedHeight;


		public static int ZoomedPointerStart = ZoomedHeightStart * 1920 * 4 + ZoomedWidthStart * 4;
		public static int ZoomedPointerEnd = ZoomedHeightEnd * 1920 * 4 + ZoomedWidthEnd * 4;
		static readonly byte[] sensorColorFrameData = new byte[1920 * 1080 * 4];
        static byte[] greenScreenFrameData = new byte[1920 * 1080 * 4];


        public unsafe static void GenerateFrame(IntPtr _ptr, int length, bool mirrored, bool zoom)
		{
            
            
            

            byte[] greenScreen = greenScreenFrameData;
			void* camData = _ptr.ToPointer();
           
            try
			{
                InitializeSensor();
                
				if (greenScreen != null)
				{

                    int colorFramePointerStart = zoom ? ZoomedPointerStart : 0;
					int colorFramePointerEnd = zoom ? ZoomedPointerEnd - 1 : greenScreen.Length - 1;
					int width = zoom ? ZoomedWidth : SensorWidth;

					if (!mirrored)
					{
						fixed (byte* sDataB = &greenScreen[colorFramePointerStart])
						fixed (byte* sDataE = &greenScreen[colorFramePointerEnd])
						{
							byte* pData = (byte*)camData;
							byte* sData = (byte*)sDataE;
							bool redo = true;

							for (; sData > sDataB;)
							{
								for (var i = 0; i < width; ++i)
								{
									var p = sData - 3;
									*pData++ = *p++;
									*pData++ = *p++;
									*pData++ = *p++;
									if (zoom)
									{
										p = sData - 3;
										*pData++ = *p++;
										*pData++ = *p++;
										*pData++ = *p++;
									}
									sData -= 4;
								}
								if (zoom)
								{
									if (redo)
									{
										sData += width * 4;
									}
									else
									{
										sData -= (SensorWidth - ZoomedWidth) * 4;
									}
									redo = !redo;

								}
							}

						}
					}
					else
					{
						fixed (byte* sDataB = &greenScreen[colorFramePointerStart])
						fixed (byte* sDataE = &greenScreen[colorFramePointerEnd])
						{
							byte* pData = (byte*)camData;
							byte* sData = (byte*)sDataE;

							var sDataBE = sData;
							var p = sData;
							var r = sData;
							bool redo = true;

							while (sData == (sDataBE = sData) &&
								   sDataB <= (sData -= (width * 4 - 1)))
							{

								r = sData;
								do
								{
									p = sData;
									*pData++ = *p++;
									*pData++ = *p++;
									*pData++ = *p++;
									if (zoom)
									{
										p = sData;
										*pData++ = *p++;
										*pData++ = *p++;
										*pData++ = *p++;
									}

								}
								while ((sData += 4) <= sDataBE);
								sData = r - 1;
								if (zoom)
								{
									if (redo)
									{
										sData += width * 4;
									}
									else
									{
										sData -= (SensorWidth - ZoomedWidth) * 4;
									}
									redo = !redo;

								}
							}
						}
					}
				}
			}
			catch
			{
                Trace.WriteLine("\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\error!!");
                byte* pData = (byte*)camData;
                for (int i = 0; i < length; ++i)
                    *pData++ = 0;
			}
		}

        private static void Reader_MultiSourceFrameArrived(object sender, MultiSourceFrameArrivedEventArgs e)
        {
            int depthWidth = 0;
            int depthHeight = 0;

            DepthFrame depthFrame = null;
            ColorFrame colorFrame = null;
            BodyIndexFrame bodyIndexFrame = null;
            bool isBitmapLocked = false;

            MultiSourceFrame multiSourceFrame = e.FrameReference.AcquireFrame();

            // If the Frame has expired by the time we process this event, return.
            if (multiSourceFrame == null)
            {
                return;
            }

            // We use a try/finally to ensure that we clean up before we exit the function.  
            // This includes calling Dispose on any Frame objects that we may have and unlocking the bitmap back buffer.
            try
            {
                depthFrame = multiSourceFrame.DepthFrameReference.AcquireFrame();
                colorFrame = multiSourceFrame.ColorFrameReference.AcquireFrame();
                bodyIndexFrame = multiSourceFrame.BodyIndexFrameReference.AcquireFrame();

                // If any frame has expired by the time we process this event, return.
                // The "finally" statement will Dispose any that are not null.
                if ((depthFrame == null) || (colorFrame == null) || (bodyIndexFrame == null))
                {
                    return;
                }









                FrameDescription colorFrameDescription = Sensor.ColorFrameSource.FrameDescription;

                int colorWidth = colorFrameDescription.Width;
                int colorHeight = colorFrameDescription.Height;

                colorMappedToDepthPoints = new DepthSpacePoint[colorWidth * colorHeight];

                bitmap = new WriteableBitmap(colorWidth, colorHeight, 96.0, 96.0, PixelFormats.Bgra32, null);

                // Calculate the WriteableBitmap back buffer size
                bitmapBackBufferSize = (uint)((bitmap.BackBufferStride * (bitmap.PixelHeight - 1)) + (bitmap.PixelWidth * bytesPerPixel));










                // Process Depth
                FrameDescription depthFrameDescription = depthFrame.FrameDescription;

                depthWidth = depthFrameDescription.Width;
                depthHeight = depthFrameDescription.Height;

                // Access the depth frame data directly via LockImageBuffer to avoid making a copy
                using (KinectBuffer depthFrameData = depthFrame.LockImageBuffer())
                {
                    coordinateMapper.MapColorFrameToDepthSpaceUsingIntPtr(
                        depthFrameData.UnderlyingBuffer,
                        depthFrameData.Size,
                        colorMappedToDepthPoints);
                }

                // We're done with the DepthFrame 
                depthFrame.Dispose();
                depthFrame = null;

                // Process Color

                
                
                // Lock the bitmap for writing
                bitmap.Lock();
                isBitmapLocked = true;

                colorFrame.CopyConvertedFrameDataToIntPtr(bitmap.BackBuffer, bitmapBackBufferSize, ColorImageFormat.Bgra);

                // We're done with the ColorFrame 
                colorFrame.Dispose();
                colorFrame = null;

                // We'll access the body index data directly to avoid a copy
                using (KinectBuffer bodyIndexData = bodyIndexFrame.LockImageBuffer())
                {
                    unsafe
                    {
                        byte* bodyIndexDataPointer = (byte*)bodyIndexData.UnderlyingBuffer;

                        int colorMappedToDepthPointCount = colorMappedToDepthPoints.Length;

                        fixed (DepthSpacePoint* colorMappedToDepthPointsPointer = colorMappedToDepthPoints)
                        {
                            // Treat the color data as 4-byte pixels
                            uint* bitmapPixelsPointer = (uint*)bitmap.BackBuffer;

                            // Loop over each row and column of the color image
                            // Zero out any pixels that don't correspond to a body index
                            for (int colorIndex = 0; colorIndex < colorMappedToDepthPointCount; ++colorIndex)
                            {
                                float colorMappedToDepthX = colorMappedToDepthPointsPointer[colorIndex].X;
                                float colorMappedToDepthY = colorMappedToDepthPointsPointer[colorIndex].Y;

                                // The sentinel value is -inf, -inf, meaning that no depth pixel corresponds to this color pixel.
                                if (!float.IsNegativeInfinity(colorMappedToDepthX) &&
                                    !float.IsNegativeInfinity(colorMappedToDepthY))
                                {
                                    // Make sure the depth pixel maps to a valid point in color space
                                    int depthX = (int)(colorMappedToDepthX + 0.5f);
                                    int depthY = (int)(colorMappedToDepthY + 0.5f);

                                    // If the point is not valid, there is no body index there.
                                    if ((depthX >= 0) && (depthX < depthWidth) && (depthY >= 0) && (depthY < depthHeight))
                                    {
                                        int depthIndex = (depthY * depthWidth) + depthX;

                                        // If we are tracking a body for the current pixel, do not zero out the pixel
                                        if (bodyIndexDataPointer[depthIndex] != 0xff)
                                        {
                                            continue;
                                        }
                                    }
                                }

                                bitmapPixelsPointer[colorIndex] = 0x0000FF00;
                            }
                        }

                        bitmap.AddDirtyRect(new System.Windows.Int32Rect(0, 0, bitmap.PixelWidth, bitmap.PixelHeight));
                    }
                }
            }
            finally
            {
                if (isBitmapLocked)
                {
                    bitmap.Unlock();
                    



                    var w = bitmap.PixelWidth;
                    var h = bitmap.PixelHeight;
                    var stride = w * ((bitmap.Format.BitsPerPixel + 7) / 8);

                    var bitmapData = new byte[h * stride];

                    bitmap.CopyPixels(bitmapData, stride, 0);

                    greenScreenFrameData = bitmapData;


                }

                if (depthFrame != null)
                {
                    depthFrame.Dispose();
                }

                if (colorFrame != null)
                {
                    colorFrame.Dispose();
                }

                if (bodyIndexFrame != null)
                {
                    bodyIndexFrame.Dispose();
                }
            }
        }
    }
}
