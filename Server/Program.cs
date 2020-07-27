using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AudioSwitcher.AudioApi.CoreAudio;

namespace PCRemote
{
    class Program
    {
        static SerialPort Port;

        static void Main(string[] args)
        {
            Port = new SerialPort("COM4", 57600, Parity.None, 8, StopBits.One);
            Port.DataReceived += Port_DataReceived;
            Port.Open();
            Application.Run();
        }

        private static void Port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string all = Port.ReadExisting();
            byte[] data = Encoding.UTF8.GetBytes(all);
            
            switch(data[0])
            {
                case 0x00://0x00 = ping so respond with 0x00 01
                    {
                        Port.Write(new byte[2] { 0x00, 0x01 }, 0, 2);
                        Console.WriteLine("Ping");
                        break;
                    }
                case 0x01://keys
                    {
                        byte result = Keys(data[1]);

                        Port.Write(new byte[2] { 0x00, result }, 0, 2);

                        break;
                    }
                case 0x02://mouse movement
                    {
                        byte result = MoveMouse(BitConverter.ToInt32(data, 1), BitConverter.ToInt32(data, 5));
                        Port.Write(new byte[2] { 0x00, result }, 0, 2);
                        break;
                    }
                case 0x03://mouse clicks
                    {
                        byte result = ClickMouse(data[1]);
                        Port.Write(new byte[2] { 0x00, result }, 0, 2);
                        break;
                    }
                case 0x04://Volume control
                    {
                        byte result = ControlVolume(data[1]);
                        Port.Write(new byte[2] { 0x00, result }, 0, 2);
                        break;
                    }
            }
        }

        #region Keys

        public static byte Keys(byte key)
        {
            keybd_event(key, 0, KeyFlags.KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
            Console.WriteLine("Pressing key " + (int)key);
            return 0x01;
        }

        [DllImport("user32.dll")]
        static extern void keybd_event(byte virtualKey, byte scanCode, KeyFlags flags, IntPtr extraInfo);

        enum KeyFlags : uint
        {
            KEYEVENTF_EXTENTEDKEY = 1,
            KEYEVENTF_KEYUP = 0
        }

        #endregion

        #region Mouse
        public static byte MoveMouse(int x, int y)
        {
            MouseOperations.SetCursorPosition(x, y);
            Console.WriteLine("Moving the mouse to " + x + ", " + y);
            return 0x01;
        }

        public static byte ClickMouse(byte button)
        {
            byte result = 0x02;

            switch(button)
            {
                case 0x00://left
                    {
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftDown);
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.LeftUp);
                        Console.WriteLine("Mouse click: left button");
                        result = 0x01;
                        break;
                    }
                case 0x01://middle
                    {
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.MiddleDown);
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.MiddleUp);
                        Console.WriteLine("Mouse click: middle button");
                        result = 0x01;
                        break;
                    }
                case 0x02://right
                    {
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.RightDown);
                        MouseOperations.MouseEvent(MouseOperations.MouseEventFlags.RightUp);
                        Console.WriteLine("Mouse click: right button");
                        result = 0x01;
                        break;
                    }
            }
            return result;
        }
        #endregion

        #region Audio
        public static byte ControlVolume(byte volume)
        {
            new CoreAudioController().DefaultPlaybackDevice.Volume = volume;
            Console.WriteLine("Setting volume to " + volume);
            return 0x01;
        }
        #endregion
    }
}
