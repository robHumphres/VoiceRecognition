using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Ports;
using System.Collections;

namespace StartingWithSpeechRecognition
{
    public class RS232
    {
        SerialPort myCOM;
        private const int SERIAL_READTIME = 1000;
        private string myPort;
        private int myBaud;
        private bool isConnected;
        private ArrayList myBuffer;
        private StringBuilder mySB;
        private bool isListening;
        private string[] splitParams;

        public RS232(string comPort, int baudRate)
        {
            myPort = comPort;
            myBaud = baudRate;
            isConnected = false;
            myBuffer = new ArrayList();
            isListening = false;
            mySB = new StringBuilder();
            splitParams = new string[1] { "\r\n" };
        }


        /// <summary>
        /// 
        /// </summary>
        public bool Connected
        {
            get { return this.isConnected; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Port
        {
            get { return this.myPort; }
            set { this.myPort = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Baud
        {
            get { return this.myBaud; }
            set { this.myBaud = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool Listening
        {
            get { return isListening; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LinesInBuffer
        {
            get { return myBuffer.Count; }
        }

        public int BytesToRead
        {
            get { return myCOM.BytesToRead; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string[] getBuffer()
        {
            //return the ArrayList
            string[] tempArray;
            tempArray = new string[myBuffer.Count];
            myBuffer.CopyTo(tempArray);
            myBuffer.Clear();
            return tempArray;
        }

        public string getData()
        {
            string temp;
            temp = mySB.ToString();
            //mySB.Clear();
            return temp;
        }

        public string DataLength
        { get { return mySB.Length.ToString(); } }


        /// <summary>
        /// 
        /// </summary>
        public void ClearBuffer()
        {
            myBuffer.Clear();
            mySB.Clear();
        }

        /// <summary>
        /// Opens the COM port specified by the constructor when this object is created. The COM port string and Baud rate integer value can be viewed or set with this objects properties.
        /// </summary>
        /// <returns>void</returns>
        public bool OpenSerialConnection()
        {

            try
            {
                myCOM = new SerialPort();
                myCOM.PortName = "COM" + myPort;
                myCOM.BaudRate = myBaud;
                myCOM.ReadTimeout = 100;
                if (myCOM.IsOpen)
                {
                    myCOM.Close();
                    myCOM.Open();
                }
                else
                    myCOM.Open();

                return this.isConnected = true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void CloseSerialConnection()
        {
            try
            {
                if (this.isListening)
                {
                    this.StopListening();
                }
                if (myCOM.IsOpen)
                {

                    myCOM.Dispose();
                    myCOM.Close();
                }
                this.isConnected = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Sends a string to the debug port and returns the strings printed in response to the command
        /// </summary>
        /// <param name="n">The string to send to the debug port</param>
        /// <returns>string[] the array of lines returned by the debug port.</returns>
        public void Write(string n)
        {
            mySB.Clear();
            try
            {
                if (myCOM.IsOpen)
                {
                    myCOM.Write(n);
                }
                else
                {
                    myCOM.Open();
                    myCOM.Write(n);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void StartListening()
        {
            try
            {
                mySB.Clear();
                myCOM.DataReceived += new SerialDataReceivedEventHandler(ImListening);
                isListening = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public void StopListening()
        {
            try
            {
                myCOM.DataReceived -= ImListening;
                isListening = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImListening(object sender, SerialDataReceivedEventArgs e)
        {

            System.DateTime timeout, endtime;
            System.TimeSpan duration;
            String tempString;

            duration = new System.TimeSpan(0, 0, 0, 0, SERIAL_READTIME);
            timeout = System.DateTime.Now;
            endtime = timeout.Add(duration);

            /*
             * Try to read for myCOM.ReadTimeout. Trap the TimeoutException
             * In the first caught timeout exception see if there are any bytes remaining to be read.
             * If there are remaining bytes read them and add them to the buffer
             * 
             */
            while (timeout < endtime)
            {
                try
                {
                    tempString = myCOM.ReadLine();
                    mySB.Append(RemoveANSITermChars(Encoding.ASCII.GetBytes(tempString)));
                    myBuffer.Add(RemoveANSITermChars(Encoding.ASCII.GetBytes(tempString)));
                }
                catch (TimeoutException)
                {
                    try
                    {
                        if (myCOM.BytesToRead > 0)
                        {
                            tempString = myCOM.ReadExisting();
                            mySB.Append(RemoveANSITermChars(Encoding.ASCII.GetBytes(tempString)));
                            myBuffer.Add(RemoveANSITermChars(Encoding.ASCII.GetBytes(tempString)));
                        }//end if
                    }
                    catch (TimeoutException) { }

                }//end catch TimeoutException
                catch (Exception) { }
                timeout = System.DateTime.Now;

            }//end while           
        }

        public static string RemoveANSITermChars(byte[] mByte)
        {
            byte[] tempBuf = new byte[mByte.Length];
            int pbuffInsert = 0, j, buffRead;
            buffRead = 0;
            if (mByte.Length > 0)
            {
                for (pbuffInsert = 0; pbuffInsert < mByte.Length; pbuffInsert++)
                {
                    if (mByte[buffRead] == 27)//this is the beginning of an escape sequence
                    {
                        for (j = buffRead + 2; j < mByte.Length; j++)
                        {
                            if ((mByte[j] > 57) && (mByte[j] < 127))//buffer[j] is a terminating char for ANSI escape sequence
                            {
                                break;
                            }

                        }
                        buffRead = j;
                        pbuffInsert--;
                    }
                    else
                    {
                        tempBuf[pbuffInsert] = mByte[buffRead];
                    }
                    buffRead++;
                    if (!(buffRead < mByte.Length))
                        break;
                }
                return Encoding.ASCII.GetString(tempBuf, 0, pbuffInsert + 1);
            }
            else
                return string.Empty;

        }//end RemoveANSITermChars
    }
}
