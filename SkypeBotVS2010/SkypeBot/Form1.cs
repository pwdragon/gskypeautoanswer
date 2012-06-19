using System;
using System.Windows.Forms;
using SKYPE4COMLib;


namespace SkypeBing
{
    public partial class frmMain : Form
    {
        private Skype skype;
        private const string trigger = "!"; // Say !help
        private const string nick = "BOT";
        
        public frmMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            skype = new Skype();
            // Use skype protocol version 7 
            skype.Attach(7, false); 
            // Listen 
            skype.MessageStatus +=new _ISkypeEvents_MessageStatusEventHandler(skype_MessageStatus);
        }
        private void skype_MessageStatus(ChatMessage msg, TChatMessageStatus status)
        {
            // Proceed only if the incoming message is a trigger
            //if (msg.Body.IndexOf(trigger) >= 0)
            if (status == TChatMessageStatus.cmsReceived)
            {
                // Remove trigger string and make lower case
                string command = msg.Body.Remove(0, trigger.Length).ToLower();

                // Send processed message back to skype chat window
                //skype.SendMessage(msg.Sender.Handle, nick + " Says: " + ProcessCommand(command));
                //skype.SendMessage(msg.Sender.Handle, "<Reposta automatica GAutoAnswer: > desculpe estou ausente deixe seu recado.");
                skype.SendMessage(msg.Sender.Handle, txtAnswer.Text);
            }
        }

        private string ProcessCommand(string str)
        {
            string result;
            switch (str)
            {
                case "hello":
                    result = "Hello!";
                    break;
                case "help":
                    result = "Sorry no help available";
                    break;
                case "date":
                    result = "Current Date is: " + DateTime.Now.ToLongDateString();
                    break;
                case "time":
                    result = "Current Time is: " + DateTime.Now.ToLongTimeString();
                    break;
                case "who":
                    result = "It is Praveen, aka NinethSense who wrote this tutorial";
                    break;
                default:
                    result = "Sorry, I do not recognize your command";
                    break;
            }

            return result;
        }


        private void ServerSimulator_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
                notifyIcon1.ShowBalloonTip(500);
                this.Hide();
            }
            else if (FormWindowState.Normal == this.WindowState)
            {
                notifyIcon1.Visible = false;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.ShowBalloonTip(1000);
            WindowState = FormWindowState.Normal;

        }
        private void executeToolStripMenuItem_Click(object sender, EventArgs e)
        {

            this.Activate();

        }



        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {

            this.Close();

        }
    }
}
