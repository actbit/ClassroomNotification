using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
namespace ClassroomNotification
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        [Flags]
        public enum PlaySoundFlags : int
        {
            SND_SYNC = 0x0000,
            SND_ASYNC = 0x0001,
            SND_NODEFAULT = 0x0002,
            SND_MEMORY = 0x0004,
            SND_LOOP = 0x0008,
            SND_NOSTOP = 0x0010,
            SND_NOWAIT = 0x00002000,
            SND_ALIAS = 0x00010000,
            SND_ALIAS_ID = 0x00110000,
            SND_FILENAME = 0x00020000,
            SND_RESOURCE = 0x00040004,
            SND_PURGE = 0x0040,
            SND_APPLICATION = 0x0080
        }
        [System.Runtime.InteropServices.DllImport("winmm.dll",
            CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern bool PlaySound(
            string pszSound, IntPtr hmod, PlaySoundFlags fdwSound);
        DispatcherTimer DispatcherTimer = new DispatcherTimer();
        public MainWindow(string title,string Code)
        {

            InitializeComponent();
            Context.Content = Code;
            Title.Content = title;
            DispatcherTimer.Interval = new TimeSpan(0,0,0,0,1200);
            DispatcherTimer.Tick += DispatcherTimer_Tick;
        }
        int size = 1000;
        private void DispatcherTimer_Tick(object sender, EventArgs e)
        {
            ;
            PlaySound("MailBeep", IntPtr.Zero,
        PlaySoundFlags.SND_ALIAS | PlaySoundFlags.SND_NODEFAULT);
            //Console.Beep(size, 1000);
            //Windowsの起動音を鳴らす
            //Windowsログオン音を鳴らす

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Random random = new Random();
            size = random.Next(500, 1500);
            DispatcherTimer.Start();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DispatcherTimer.Stop();

            this.Close();
        }
    }
}
