using Google.Apis.Auth.OAuth2;
using Google.Apis.Classroom.v1;
using Google.Apis.Classroom.v1.Data;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Windows;
using System.Xml.Serialization;

namespace ClassroomNotification
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App :  Application
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
        static string[] Scopes = {
            ClassroomService.Scope.ClassroomAnnouncements,
            ClassroomService.Scope.ClassroomCourses,
            ClassroomService.Scope.ClassroomCourseworkmaterials,
            ClassroomService.Scope.ClassroomCourseworkMe,
            ClassroomService.Scope.ClassroomCourseworkStudents,
            ClassroomService.Scope.ClassroomGuardianlinksStudents,
            ClassroomService.Scope.ClassroomProfileEmails,
            ClassroomService.Scope.ClassroomProfilePhotos,
            ClassroomService.Scope.ClassroomPushNotifications,
            ClassroomService.Scope.ClassroomRosters,
            ClassroomService.Scope.ClassroomTopics,

        };
        private DispatcherTimer _timer = new DispatcherTimer();
        static string ApplicationName = "ClassroomNotification";
        List<CoursesResource.CourseWorkResource.ListRequest> CourseWorkListRequests=new List<CoursesResource.CourseWorkResource.ListRequest>();
        List<CoursesResource.CourseWorkMaterialsResource.ListRequest> CourseWorkMaterialsListRequests=new List<CoursesResource.CourseWorkMaterialsResource.ListRequest>();
        List<CoursesResource.TopicsResource.ListRequest> TopicsResourceListRequests=new List<CoursesResource.TopicsResource.ListRequest>();
        List<CoursesResource.AliasesResource.ListRequest> AliasesResourceListRequests=new List<CoursesResource.AliasesResource.ListRequest>();
        List<CoursesResource.AnnouncementsResource.ListRequest> AnnouncementsResourceListRequests=new List<CoursesResource.AnnouncementsResource.ListRequest>();
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            if (File.Exists("Save.xml"))
            {
                SaveData=XMLClass.LoadFromFile<SaveData>(File.ReadAllText("Save.xml"));
            }
            else
            {
                SaveData =new SaveData();

                File.WriteAllText("Save.xml", XMLClass.SaveToFile(SaveData));
            }
            Microsoft.Win32.RegistryKey regkey2 =
    Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
    @"Software\Microsoft\Windows\CurrentVersion\Run", false);
            //値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
            var d = regkey2.GetValue(System.Windows.Forms.Application.ProductName);
            //閉じる
            regkey2.Close();

            if (d == null||d.ToString()!= System.Windows.Forms.Application.ExecutablePath)
            {
                //Runキーを開く
                Microsoft.Win32.RegistryKey regkey =
                    Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Run", true);
                //値の名前に製品名、値のデータに実行ファイルのパスを指定し、書き込む
                regkey.SetValue(System.Windows.Forms.Application.ProductName, System.Windows.Forms.Application.ExecutablePath);
                //閉じる
                regkey.Close();
            }

            _timer.Interval = new TimeSpan(0, 1, 30);
            _timer.Tick += _timer_Tick;
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Classroom API service.
            var service = new ClassroomService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.

            CoursesResource.ListRequest request = service.Courses.List();
            request.PageSize = 10;

            //CloudPubsubTopic
            // List courses.
            ListCoursesResponse response = request.Execute();
            Console.WriteLine("Courses:");
            if (response.Courses != null && response.Courses.Count > 0)
            {
                foreach (var course in response.Courses)
                {
                    CourseWorkListRequests.Add(service.Courses.CourseWork.List(course.Id));
                    CourseWorkListRequests[CourseWorkListRequests.Count-1].PageSize = 10;
                    CourseWorkMaterialsListRequests.Add(service.Courses.CourseWorkMaterials.List(course.Id));
                    CourseWorkMaterialsListRequests[CourseWorkMaterialsListRequests.Count-1].PageSize = 10;
                    TopicsResourceListRequests.Add(service.Courses.Topics.List(course.Id));
                    TopicsResourceListRequests[TopicsResourceListRequests.Count-1].PageSize = 10;
                    AliasesResourceListRequests.Add(service.Courses.Aliases.List(course.Id));
                    AliasesResourceListRequests[AliasesResourceListRequests.Count-1].PageSize = 10;
                    AnnouncementsResourceListRequests.Add(service.Courses.Announcements.List(course.Id));
                    AnnouncementsResourceListRequests[AnnouncementsResourceListRequests.Count-1].PageSize = 10;
                    
                    Console.WriteLine("{0} ({1})", course.Name, course.Id);
                }
            }
            else
            {
                Console.WriteLine("No courses found.");
            }
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
            _timer.Start();
        }
        SaveData SaveData = new SaveData();


        private async void _timer_Tick(object sender, EventArgs e)
        {
            foreach(var Request in CourseWorkListRequests)
            {
                var r1 = await Request.ExecuteAsync();
                foreach(var t in r1.CourseWork)
                {
                    if ((DateTime)t.UpdateTime > SaveData.CourseWorkDateTime)
                    {
                        MainWindow mainWindow = new MainWindow(t.Title, t.Description);
                        mainWindow.Show();
                    }
                    
                }

            }

            SaveData.CourseWorkDateTime = DateTime.Now;
            File.WriteAllText("Save.xml", XMLClass.SaveToFile(SaveData));

            foreach (var Request in CourseWorkMaterialsListRequests)
            {
                var r1 = await Request.ExecuteAsync();
                foreach (var t in r1.CourseWorkMaterial)
                {
                    if ((DateTime)t.UpdateTime > SaveData.CourseWorkMaterialsDateTime)
                    {
                        MainWindow mainWindow = new MainWindow(t.Title, t.Description);
                        mainWindow.Show();
                    }

                }
            }
            SaveData.CourseWorkMaterialsDateTime = DateTime.Now;
            File.WriteAllText("Save.xml", XMLClass.SaveToFile(SaveData));

            foreach (var Request in TopicsResourceListRequests)
            {
                var r1 = await Request.ExecuteAsync();
                foreach (var t in r1.Topic)
                {
                    if ((DateTime)t.UpdateTime > SaveData.TopicsResourceDateTime)
                    {
                        MainWindow mainWindow = new MainWindow("", t.Name);
                        mainWindow.Show();
                    }
                    
                }
            }
            SaveData.TopicsResourceDateTime = DateTime.Now;
            File.WriteAllText("Save.xml", XMLClass.SaveToFile(SaveData));

            foreach (var Request in AnnouncementsResourceListRequests)
            {
                var r1 = await Request.ExecuteAsync();

                foreach (var t in r1.Announcements)
                {
                    if ((DateTime)t.UpdateTime > SaveData.AnnouncementsResourceDateTime)
                    {
                        MainWindow mainWindow = new MainWindow("", t.Text);
                        mainWindow.Show();
                    }

                }

            }
            SaveData.AnnouncementsResourceDateTime = DateTime.Now;
            File.WriteAllText("Save.xml", XMLClass.SaveToFile(SaveData));

        }
    }
    public class SaveData
    {
        public DateTime CourseWorkDateTime = DateTime.Now;
        public DateTime CourseWorkMaterialsDateTime = DateTime.Now;
        public DateTime TopicsResourceDateTime = DateTime.Now;
        public DateTime AnnouncementsResourceDateTime = DateTime.Now;
    }
    static class XMLClass
    {
        public static string SaveToFile<T>(T control)
        {
            var writer = new StringWriter(); // 出力先のWriterを定義
            var serializer = new XmlSerializer(typeof(T));
            serializer.Serialize(writer, control);

            var xml = writer.ToString();
            //Console.WriteLine(xml);


            return xml;
        }

        public static T LoadFromFile<T>(string s)
        {
            var serializer = new XmlSerializer(typeof(T));
            var deserializedBook = (T)serializer.Deserialize(new StringReader(s));
            return deserializedBook;
        }
    }
}
