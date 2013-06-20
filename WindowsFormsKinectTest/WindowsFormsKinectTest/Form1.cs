using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using Microsoft.Kinect;
using Microsoft.Kinect.Toolkit;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.AudioFormat;
using System.Runtime.InteropServices;

namespace WindowsFormsKinectTest
{
    public partial class Form1 : Form
    {
        private KinectSensorChooser _chooser;
        private Bitmap _bitmap;

        public int PlayerId;

        private SpeechRecognitionEngine speechEngine;

        public Form1()
        {
            InitializeComponent();
        }

        private static RecognizerInfo GetKinectRecognizer()
        {
            foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "en-US".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return recognizer;
                }
            }
            return null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _chooser = new KinectSensorChooser();
            _chooser.KinectChanged += ChooserSensorChanged;
            _chooser.Start();
            KinectSensor kinect = _chooser.Kinect;
            if(kinect == null)
                MessageBox.Show("No Kinect connected!", "No Kinect error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            this.Close();
        }

        private void Form1_Close(object sender, EventArgs e)
        {
            if (null != this._chooser.Kinect)
            {
                this._chooser.Kinect.AudioSource.Stop();

                this._chooser.Kinect.Stop();
            }

            if (null != this.speechEngine)
            {
                this.speechEngine.SpeechRecognized -= SpeechRecognized;
                this.speechEngine.SpeechRecognitionRejected -= SpeechRecognitionRejected;
                this.speechEngine.RecognizeAsyncStop();
            }
        }

        void ChooserSensorChanged(object sender, KinectChangedEventArgs e)
        {
            var old = e.OldSensor;
            StopKinect(old);

            var newsensor = e.NewSensor;
            if (newsensor == null)
            {
                return;
            }

            newsensor.SkeletonStream.Enable();
            newsensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
            newsensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);
            newsensor.AllFramesReady += SensorAllFramesReady;

            try
            {
                newsensor.Start();
                rtbMessages.Text = "Kinect Started" + "\r";

                RecognizerInfo ri = GetKinectRecognizer();
                if (null != ri)
                {
                    this.speechEngine = new SpeechRecognitionEngine(ri.Id);

                    // Create a simple grammar that recognizes "red", "green", or "blue".
                    Choices colors = new Choices();
                    colors.Add(new SemanticResultValue("red", "RED"));
                    colors.Add(new SemanticResultValue("green", "GREEN"));
                    colors.Add(new SemanticResultValue("blue", "BLUE"));

                    // Create a GrammarBuilder object and append the Choices object.
                    var gb = new GrammarBuilder { Culture = ri.Culture };
                    gb.Append(colors);

                    // Create the Grammar instance and load it into the speech recognition engine.
                    var g = new Grammar(gb);

                    speechEngine.LoadGrammar(g);

                    // Register a handler for the SpeechRecognized event.
                    speechEngine.SpeechRecognized += SpeechRecognized;

                    speechEngine.SpeechRecognitionRejected += SpeechRecognitionRejected;

                    // Configure the input to the recognizer.
                    // Start recognition.
                    speechEngine.SetInputToAudioStream(
                        newsensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                    speechEngine.RecognizeAsync(RecognizeMode.Multiple);

                    rtbMessages.AppendText("Speech recognition started\r");
                    rtbMessages.ScrollToCaret();
                }
            }
            catch (System.IO.IOException)
            {
                rtbMessages.Text = "Kinect Not Started" + "\r";
                //maybe another app is using Kinect
                _chooser.TryResolveConflict();
            }
        }

        private void StopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                if (sensor.IsRunning)
                {
                    sensor.Stop();
                    sensor.AudioSource.Stop();
                }
            }
        }

        void SensorAllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            SensorDepthFrameReady(e);
            SensorSkeletonFrameReady(e);
            video.Image = _bitmap;
        }

        void SensorSkeletonFrameReady(AllFramesReadyEventArgs e)
        {
            using (SkeletonFrame skeletonFrameData = e.OpenSkeletonFrame())
            {
                if (skeletonFrameData == null)
                {
                    return;
                }

                var allSkeletons = new Skeleton[skeletonFrameData.SkeletonArrayLength];
                skeletonFrameData.CopySkeletonDataTo(allSkeletons);

                foreach (Skeleton sd in allSkeletons)
                {
                    // If this skeleton is no longer being tracked, skip it
                    if (sd.TrackingState != SkeletonTrackingState.Tracked)
                    {
                        continue;
                    }
                    
                    PlayerId = sd.TrackingId;

                    if (_bitmap != null)
                        _bitmap = AddSkeletonToDepthBitmap(sd, _bitmap, true);

                }
            }
        }

        /// <summary>
        /// This method draws the joint dots and skeleton on the depth image of the depth display
        /// </summary>
        /// <param name="skeleton"></param>
        /// <param name="bitmap"></param>
        /// <param name="isActive"> </param>
        /// <returns></returns>
        private Bitmap AddSkeletonToDepthBitmap(Skeleton skeleton, Bitmap bitmap, bool isActive)
        {
            Pen pen;
            Pen handPen;

            var gobject = Graphics.FromImage(bitmap);

            if (isActive)
            {
                pen = new Pen(Color.Green, 5);
                handPen = new Pen(Color.Red, 5);
            }
            else
            {
                pen = new Pen(Color.DeepSkyBlue, 5);
                handPen = new Pen(Color.DeepSkyBlue, 5);
            }

            var head = CalculateJointPosition(bitmap, skeleton.Joints[JointType.Head]);
            var neck = CalculateJointPosition(bitmap, skeleton.Joints[JointType.ShoulderCenter]);
            var rightshoulder = CalculateJointPosition(bitmap, skeleton.Joints[JointType.ShoulderRight]);
            var leftshoulder = CalculateJointPosition(bitmap, skeleton.Joints[JointType.ShoulderLeft]);
            var rightelbow = CalculateJointPosition(bitmap, skeleton.Joints[JointType.ElbowRight]);
            var leftelbow = CalculateJointPosition(bitmap, skeleton.Joints[JointType.ElbowLeft]);
            var rightwrist = CalculateJointPosition(bitmap, skeleton.Joints[JointType.WristRight]);
            var leftwrist = CalculateJointPosition(bitmap, skeleton.Joints[JointType.WristLeft]);

            label6.Text = "right elbow X = " + rightelbow.X;
            label7.Text = "right elbow Y = " + rightelbow.Y;
            label8.Text = "right elbow Z = " + rightelbow.Z;

            label2.Text = "right wrist X = " + rightwrist.X;
            label3.Text = "right wrist Y = " + rightwrist.Y;
            label4.Text = "right wrist Z = " + rightwrist.Z;

            //var spine = CalculateJointPosition(bitmap, skeleton.Joints[JointType.Spine]);
            var hipcenter = CalculateJointPosition(bitmap, skeleton.Joints[JointType.HipCenter]);
            var hipleft = CalculateJointPosition(bitmap, skeleton.Joints[JointType.HipLeft]);
            var hipright = CalculateJointPosition(bitmap, skeleton.Joints[JointType.HipRight]);
            var kneeleft = CalculateJointPosition(bitmap, skeleton.Joints[JointType.KneeLeft]);
            var kneeright = CalculateJointPosition(bitmap, skeleton.Joints[JointType.KneeRight]);
            var ankleleft = CalculateJointPosition(bitmap, skeleton.Joints[JointType.AnkleLeft]);
            var ankleright = CalculateJointPosition(bitmap, skeleton.Joints[JointType.AnkleRight]);

            gobject.DrawEllipse(pen, new Rectangle((int)head.X - 25, (int)head.Y - 25, 50, 50));
            gobject.DrawEllipse(pen, new Rectangle((int)neck.X - 5, (int)neck.Y, 10, 10));
            gobject.DrawLine(pen, head.X, head.Y + 25, neck.X, neck.Y);

            gobject.DrawLine(pen, neck.X, neck.Y, rightshoulder.X, rightshoulder.Y);
            gobject.DrawLine(pen, neck.X, neck.Y, leftshoulder.X, leftshoulder.Y);
            gobject.DrawLine(pen, rightshoulder.X, rightshoulder.Y, rightelbow.X, rightelbow.Y);
            gobject.DrawLine(pen, leftshoulder.X, leftshoulder.Y, leftelbow.X, leftelbow.Y);

            gobject.DrawLine(pen, rightshoulder.X, rightshoulder.Y, hipcenter.X, hipcenter.Y);
            gobject.DrawLine(pen, leftshoulder.X, leftshoulder.Y, hipcenter.X, hipcenter.Y);

            gobject.DrawEllipse(handPen, new Rectangle((int)rightwrist.X - 10, (int)rightwrist.Y - 10, 20, 20));
            gobject.DrawLine(handPen, rightelbow.X, rightelbow.Y, rightwrist.X, rightwrist.Y);
            gobject.DrawEllipse(pen, new Rectangle((int)leftwrist.X - 10, (int)leftwrist.Y - 10, 20, 20));
            gobject.DrawLine(pen, leftelbow.X, leftelbow.Y, leftwrist.X, leftwrist.Y);

            gobject.DrawLine(pen, hipcenter.X, hipcenter.Y, hipleft.X, hipleft.Y);
            gobject.DrawLine(pen, hipcenter.X, hipcenter.Y, hipright.X, hipright.Y);
            gobject.DrawLine(pen, hipleft.X, hipleft.Y, kneeleft.X, kneeleft.Y);
            gobject.DrawLine(pen, hipright.X, hipright.Y, kneeright.X, kneeright.Y);
            gobject.DrawLine(pen, kneeright.X, kneeright.Y, ankleright.X, ankleright.Y);
            gobject.DrawLine(pen, kneeleft.X, kneeleft.Y, ankleleft.X, ankleleft.Y);

            /**
             * Pointing vector - rigth hand from elbow to wrist
            /*
            var vecX = Math.Abs(rightelbow.X - rightwrist.X);
            var vecY = Math.Abs(rightelbow.Y - rightwrist.Y);
            var vecZ = Math.Abs(rightelbow.Z - rightwrist.Z);
            
            vecLen = Math.Sqrt(Math.Pow(vecX, 2) + Math.Pow(vecY, 2) + Math.Pow(vecZ, 2));
            */

            return bitmap;
        }

        /// <summary>
        /// This method turns a skeleton joint position vector into one that is scaled to the depth image
        /// </summary>
        /// <param name="bitmap"></param>
        /// <param name="joint"></param>
        /// <returns></returns>
        protected SkeletonPoint CalculateJointPosition(Bitmap bitmap, Joint joint)
        {
            var jointx = joint.Position.X;
            var jointy = joint.Position.Y;
            var jointz = joint.Position.Z;

            if (jointz < 1)
                jointz = 1;

            var jointnormx = jointx / (jointz * 1.1f);
            var jointnormy = -(jointy / jointz * 1.1f);
            var point = new SkeletonPoint();
            point.X = (jointnormx + 0.5f) * bitmap.Width;
            point.Y = (jointnormy + 0.5f) * bitmap.Height;
            point.Z = jointz;
            return point;
        }

        void SensorDepthFrameReady(AllFramesReadyEventArgs e)
        {
            // if the window is displayed, show the depth buffer image
            if (WindowState != FormWindowState.Minimized)
            {
                using (var frame = e.OpenDepthImageFrame())
                {
                    _bitmap = CreateBitMapFromDepthFrame(frame);
                }
            }
        }

        private Bitmap CreateBitMapFromDepthFrame(DepthImageFrame frame)
        {
            if (frame != null)
            {
                var bitmapImage = new Bitmap(frame.Width, frame.Height, PixelFormat.Format16bppRgb565);

                //Copy the depth frame data onto the bitmap  
                var _pixelData = new short[frame.PixelDataLength];
                frame.CopyPixelDataTo(_pixelData);
                BitmapData bmapdata = bitmapImage.LockBits(new Rectangle(0, 0, frame.Width,
                 frame.Height), ImageLockMode.WriteOnly, bitmapImage.PixelFormat);
                IntPtr ptr = bmapdata.Scan0;
                Marshal.Copy(_pixelData, 0, ptr, frame.Width * frame.Height);
                bitmapImage.UnlockBits(bmapdata);

                return bitmapImage;
            }
            return null;
        }

        // Create a simple handler for the SpeechRecognized event.
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // Speech utterance confidence below which we treat speech as if it hadn't been heard
            const double ConfidenceThreshold = 0.3;

            if (e.Result.Confidence >= ConfidenceThreshold)
            {
                switch (e.Result.Semantics.Value.ToString())
                {
                    case "RED":
                        rtbMessages.AppendText("Speech recognized: red\r");
                        rtbMessages.ScrollToCaret();
                        break;

                    case "GREEN":
                        rtbMessages.AppendText("Speech recognized: green\r");
                        rtbMessages.ScrollToCaret();
                        break;

                    case "BLUE":
                        rtbMessages.AppendText("Speech recognized: blue\r");
                        rtbMessages.ScrollToCaret();
                        break;
                }
            }
            else
            {
                rtbMessages.AppendText("Speech not recognized\r");
                rtbMessages.ScrollToCaret();
            }
        }

        // Create a simple handler for the SpeechRejected event.
        private void SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            rtbMessages.AppendText("Speech rejected\r");
            rtbMessages.ScrollToCaret();
        }

    }
}