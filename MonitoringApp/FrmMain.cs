using System;
using System.IO;
using System.Windows.Forms;
using WndList = System.Collections.Generic.List<System.IntPtr>;

namespace MonitoringApp
{
    public partial class FrmMain : Form
    {
        System.Timers.Timer Timer;
        
        public FrmMain()
        {
            InitializeComponent();



            Timer = new System.Timers.Timer();
            Timer.Elapsed += TimeExport_Elapsed;
            Timer.AutoReset = false;
            Timer.Enabled = true;
            Timer.Stop();
            string[] args = Environment.GetCommandLineArgs();
             
            if (args.Length>1 && args[1] == "1")
            {

                btnstart_Click(null, null);
            }


        }
        void TimeExport_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {

            try
            {

                Timer.Stop();

                DoJob();


            }
            catch (Exception ex)
            {
                Mail.SendMail("MonitoringApp - Errore", ex.Message);

            }

            finally
            {
                Timer.Start();

            }
        }
        private void DoJob()
        {
            try
            {
                var folderscreen = "screenshot";
                var MailObj = System.Configuration.ConfigurationManager.AppSettings["MailObj"].ToString();
                var MailBody = System.Configuration.ConfigurationManager.AppSettings["MailBody"].ToString();
                var filename = $@"Screenshot_{DateTime.Now.ToString("yyyyMMddHHmmss")}.jpg";
                string folder = System.Windows.Forms.Application.StartupPath + "\\";
                if (!Directory.Exists(folder+ folderscreen))
                {
                    Directory.CreateDirectory(folder + folderscreen);
                }
                ScreenMonitorLib.SnapShot snp = new ScreenMonitorLib.SnapShot(folder, folderscreen + "\\" + filename); 
                IntPtr hDesktop = IntPtr.Zero;
                WndList lst = snp.GetDesktopWindows(hDesktop);
                if (lst.Count > 0)
                {
                    snp.SaveAllSnapShots(lst);

                    Mail.SendMail(MailObj, MailBody , folder + folderscreen + "\\" + filename);
                }
                else
                {
                    Mail.SendMail("MonitoringApp-Error", "Nessun processo attivo");
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        private void btnstart_Click(object sender, EventArgs e)
        {
            try
            {
                var intervallo = System.Configuration.ConfigurationManager.AppSettings["IntervalloTimer"].ToString();
                var type = intervallo.Substring(0, 1);
                if (type == "M")
                {
                    Timer.Interval = (1000 * 60) * Convert.ToDouble(intervallo.Substring(1));
                }
                else
                {
                    Timer.Interval = (1000 * 60 * 60) * Convert.ToDouble(intervallo.Substring(1));
                }
                
                Timer.Start();
                btnstart.Enabled = false;
            }
            catch (Exception)
            {

                MessageBox.Show("Attenzione,Formato intervallo non valido");

            }
            

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Timer.Stop();
                this.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Errore: "  + ex.Message);
            }
            
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            try
            {

                btnTest.Enabled = false;

                DoJob();


            }
            catch (Exception ex)
            {

                MessageBox.Show("Errore: "+ ex.Message);

            }

            finally
            {
                btnTest.Enabled = true;
            }
        }
    }
}
