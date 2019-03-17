using System;
using System.Device;
using System.Management;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using Microsoft.Speech.AudioFormat;
using System.Windows.Forms;
using NAudio;
using NAudio.Wave;
using System.Media;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;

namespace SpechRecognitionConsole
{
    class Program
    {
        //Для сворачивания
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);
        //Скриншот
        


        static StreamReader SR = new StreamReader(@"D:\Text.txt", System.Text.Encoding.UTF8);
        static string Larisa = "Лариса";
        static object obj = new object();
        static StreamWriter SW = new StreamWriter("D:\\Recognition.txt", true);
        static SpeechSynthesizer synth = new SpeechSynthesizer();
        static SpeechSynthesizer synth1 = new SpeechSynthesizer();
        static DirectoryInfo dir;
        
        AudioFileReader audioFileReader = new AudioFileReader(@"D:\\ChillingMusic.wav");        

        static SoundPlayer m_SoundPlayer;
        static Stream OutputStream;
        
        static void Main(string[] args)
        {
            Console.SetWindowSize(85, 20);
            dir = new DirectoryInfo(@"D:\\Synthesizer.wav");
            synth.SetOutputToWaveFile(@"D:\\Synthesizer.wav");            
            synth.Rate = 2;
            synth.Volume = 80;
            //synth.SelectVoiceByHints(VoiceGender gender, VoiceAge age, int voiceAlternate);
            //synth.SelectVoiceByHints(VoiceGender.Male, VoiceAge.Child, 1);
            synth.SetOutputToDefaultAudioDevice();
            //synth.SetOutputToAudioStream(OutputStream, new SpeechAudioFormatInfo(32000, AudioBitsPerSample.Sixteen, AudioChannel.Mono));
            //synth.SetOutputToWaveStream(OutputStream);
            //m_SoundPlayer = new System.Media.SoundPlayer(OutputStream);

            m_SoundPlayer = new SoundPlayer(@"D:\\ChillingMusic.wav");

            //synth1.SetOutputToWaveFile(@"D:\\Synthesizer1.wav");
            //synth1.Rate = 2;
            //synth1.Volume = 80;
            //synth1.SetOutputToDefaultAudioDevice();

            System.Globalization.CultureInfo russianCulture = new System.Globalization.CultureInfo("ru-ru");    //Культурные данные русского языка
            
            SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine(russianCulture);                   //Привязываем культурные данные русского языка к распознавателю
                      
            Choices numbers = new Choices();                                                                    //Список слов для распознавания
            //string[] commands = new string[] { "Выключись", "Пути",
            //                            "Консоль", "Хром",
            //                            "Нотпад", "Кодесис",
            //                            "Пинг таблицы",
            //                            "Сетевые подключения",
            //                            "Проверка Сириусов",
            //                            "Проверка Интеров",
            //                            "Удаленное управление",
            //                            "Интер", "Старт",
            //                            "Студия", "Привет",
            //                            "Как дела", "Детка",
            //                            "Молодец", "Скажи диме привет",
            //                            "Скажи Ване привет", "Да",
            //                            "Почта", "3 закона робототехники",
            //                            "Детка", "История", "Время",
            //                            "Кто твой хозяин", "Давай детка",
            //                            "Скриншот", "Калькулятор",
            //                            "Громкость звука", "Погода",
            //                            "Подсветка", "Отбой", "Проводник",
            //                            "Свернуть все", "Закрыть вкладку", "Следующая вкладка",
            //                            "Пауза", "Дальше", "Дальше дальше", "Дальше дальше дальше", "На весь экран",
            //                            "Следующее видео", "Выключить звук", "Извлечь флешку"
            //                            };
             string[] commands = new string[] { "Выключись", "Пути",
                                        "Консоль", "Хром",
                                        "Нотпад", "Кодесис",
                                        "Пинг таблицы",
                                        "Сетевые подключения",
                                        "Проверка Сириусов",
                                        "Проверка Интеров",
                                        "Удаленное управление",
                                        "Интер", "Старт",
                                        "Студия",
                                        "Время",                                        
                                        "Скриншот", "Калькулятор",
                                        "Громкость звука", "Погода",
                                        "Подсветка", "Закрыть", "Проводник",
                                        "Свернуть все", "Блокировка",
                                        "Читать"
                                        };

            numbers.Add(commands);
            Console.WriteLine("\t\t\t\tДоступные команды");
            int index1 = 0;
            ConsoleColor old = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            foreach (string command in commands)
            {
                index1++;
                Console.Write("{0,25}, ",command);
                
                if (index1 == 3)
                {
                    index1 = 0;
                    Console.WriteLine();
                }                    
            }
            Console.ForegroundColor = old;

            GrammarBuilder gb = new GrammarBuilder();                                                           //Объект GrammarBuilder
            gb.Culture = russianCulture;
            gb.Append(numbers);

            Grammar g = new Grammar(gb);
            recognizer.LoadGrammar(g);              

            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);    //Привязываем к событию метод, который будет 
                        
            recognizer.SetInputToDefaultAudioDevice();                                                          //Выбираем микрофон по умолчанию
                        
            recognizer.RecognizeAsync(RecognizeMode.Multiple);                                                  //Асинхронный старт распознавания
            DateTime NowTime = DateTime.Now;
            int Hour = 22;
            int Minute = 17;

            //if (Console.ReadLine()=="L")
            //{
            //    Console.Beep(1500, 200);                                                                                                //Мелодия марио
            //    Console.Beep(1500, 200);
            //    Voice("Извлекаю флешку");
            //    foreach (Volume device in volumeDeviceClass.Devices)
            //    {
            //        // is this volume on USB disks?
            //        if (!device.IsUsb)
            //            continue;

            //        // is this volume a logical disk?
            //        if ((device.LogicalDrive == null) || (device.LogicalDrive.Length == 0))
            //            continue;

            //        device.Eject(true); // allow Windows to display any relevant UI
            //    }
            //}
            while (true)                        //Бесконечный цикл, для того чтобы консоль не закрылась, каждую 0 минуту и 0 секунду будет говориться время
            {
                NowTime = DateTime.Now;
                if((int)NowTime.Minute == 0 && (int)NowTime.Second == 0)
                {
                    TimeSpeak(NowTime, "");
                    if (NowTime.Hour == 18 && NowTime.Minute == 0) TimeSpeak(NowTime, "Пора домой");
                    else if (NowTime.Hour == 12 && NowTime.Minute == 0) TimeSpeak(NowTime, "Пора обедать");                    
                }
                Thread.Sleep(500);
            }
            Console.ReadLine();
            SR.Close();            
        }
                
        static void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)                     //Метод выполняемый каждый раз при распознавании
        {

            RecordEntry(e.Result.Text);
            if (e.Result.Confidence > 0.6)                                                                        //Если совпадение более 70% то
            {
                //Console.WriteLine("Recognized text: " + e.Result.Text);
                if (e.Result.Text =="Выключись")                                                               //Выполняем команду выключить
                {
                    Console.Beep(1500, 200);
                    //PlayMario();                                                                                //Мелодия марио
                    Console.Beep(1500, 200);
                    RecordEntry(e.Result.Text);
                    //Process.Start("ShutDown", "/s -t 00");
                    Voice("Выключаюсь");
                    Thread.Sleep(2000);
                    Voice("Я Шучу");

                    //Process.Start("shutdown.exe", "/s /f /t 0");
                }
                if (e.Result.Text == "Блокировка")                                                               //Выполняем команду выключить
                {
                    Console.Beep(1500, 200);
                    //PlayMario();                                                                                //Мелодия марио
                    Console.Beep(1500, 200);
                    Voice("Режим блокировки");
                    //SendKeys.SendWait("(^{ESC}L)");             //Win + L = Ctrl + Esc + L
                    Process.Start(@"C:\WINDOWS\system32\rundll32.exe", "user32.dll,LockWorkStation");

                    //Process.Start("shutdown.exe", "/s /f /t 0");
                }
                //if (e.Result.Text == "Извлечь флешку")                                                               //Выполняем команду выключить
                //{
                //    Console.Beep(1500, 200);                                                                                                //Мелодия марио
                //    Console.Beep(1500, 200);
                //    Voice("Извлекаю флешку");
                //    foreach (Volume device in volumeDeviceClass.Devices)
                //    {
                //        // is this volume on USB disks?
                //        if (!device.IsUsb)
                //            continue;

                //        // is this volume a logical disk?
                //        if ((device.LogicalDrive == null) || (device.LogicalDrive.Length == 0))
                //            continue;

                //        device.Eject(true); // allow Windows to display any relevant UI
                //    }

                //}
                if (e.Result.Text == "Пути")
                {
                    ProcStart(@"C:\Program Files\PuTTY\putty.exe", e.Result.Text);
                }
                if (e.Result.Text == "Консоль")
                {
                    ProcStart(@"C:\Windows\system32\cmd.exe", e.Result.Text);
                }
                if (e.Result.Text == "Хром")
                {
                    //Process.Start("http://google.com");
                    ProcStart(@"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", e.Result.Text);
                }
                if (e.Result.Text == "Нотпад")
                {                    
                    ProcStart(@"C:\Program Files (x86)\Notepad++\notepad++.exe", e.Result.Text);
                }
                if (e.Result.Text == "Кодесис")
                {                    
                    ProcStart(@"C:\Program Files (x86)\3S CODESYS\CODESYS\Common\CODESYS.exe", e.Result.Text);
                }
                if (e.Result.Text == "Пинг таблицы")
                {
                    ProcStart(@"D:\C# projects\RedMachine.exe", e.Result.Text);
                }
                if (e.Result.Text == "Сетевые подключения")
                {
                    ProcStart(@"::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", e.Result.Text);
                }
                if (e.Result.Text == "Проверка Сириусов")
                {
                    ProcStart(@"D:\Работа\Проекты\ТПК\Э\Проверка Сириусов по станциям", e.Result.Text);
                }
                if (e.Result.Text == "Проверка Интеров")
                {
                    ProcStart(@"D:\Работа\Проекты\ТПК\Э\Проверка ИнТер-825", e.Result.Text);
                }
                if (e.Result.Text == "Удаленное управление")
                {
                    ProcStart(@"C:\Windows\system32\mstsc.exe", e.Result.Text);
                }
                if (e.Result.Text == "Старт")
                {
                    ProcStart(@"C:\Program Files(x86)\Radius\Start3\Start3.exe", e.Result.Text);
                }
                if (e.Result.Text == "Интер")
                {
                    ProcStart(@"D:\Работа\Проекты\ТПК\Э\Проверка ИнТер-825\ПО ИНТЕР (5.02.18)\ИнТер-825(Москва)\АСУ ИнТер-825(Москва)\Inter08.exe", e.Result.Text);
                }
                if (e.Result.Text == "Студия")
                {
                    ProcStart(@"C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\Common7\IDE\devenv.exe", e.Result.Text);
                }
                if (e.Result.Text == "Почта")
                {
                    synth.Speak("");
                }                             
                if (e.Result.Text == "Время")
                {                    
                    DateTime NowTime = DateTime.Now;
                    TimeSpeak(NowTime);                    
                }          
                
                if (e.Result.Text == "Детка")
                {
                    synth.Speak("Я вас не подведу мой Господин");
                }
                //Громкость звука
                if (e.Result.Text == "Громкость звука")
                {
                    synth.Speak("Уменьшила звук");
                    if(synth.Volume>=10)
                    {
                        synth.Volume -= 20;
                    }                    
                }                
                if (e.Result.Text == "Калькулятор")
                {
                    ProcStart(@"calc.exe", e.Result.Text);
                }
                if (e.Result.Text == "Погода")
                {
                    ProcStart(@"https://yandex.ru/pogoda/moscow?from=serp_title", "Погоду");
                }
                //Скриншот
                if (e.Result.Text == "Скриншот")
                {
                    synth.Speak("Сохраняю экран");
                    ScreenShot();
                }
                //Отбой
                if (e.Result.Text == "Закрыть")           //ALt + F4
                {
                    synth.Speak("Закрываю");
                    SendKeys.SendWait("(%{F4})");
                }
                if (e.Result.Text == "Проводник")
                {
                    ProcStart(@"explorer.exe", e.Result.Text);
                }
                //SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL, IntPtr.Zero); 
                if (e.Result.Text == "Свернуть все")
                {
                    synth.Speak("Сворачиваю");
                    const int WM_COMMAND = 0x111;
                    const int MIN_ALL = 419;
                    IntPtr lHwnd = FindWindow("Shell_TrayWnd", null);
                    SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL, IntPtr.Zero);                    
                }   
                if(e.Result.Text == "Читать")
                {
                    if(Clipboard.ContainsText() == true)
                    {
                        synth.Speak("Читаю");
                        string BufText = Clipboard.GetText();
                        HistoryReader(BufText);
                    }
                }
            }
        }
        static void TimeSpeak(DateTime NowTime, string str = "")
        {
            string HourWord;
            string MinetsWord;
            if (NowTime.Hour == 1) HourWord = "час";
            else if (NowTime.Hour > 1 && NowTime.Hour < 5) HourWord = "часа";
            else if ((NowTime.Hour > 4 && NowTime.Hour < 21) || NowTime.Hour == 0) HourWord = "часов";
            else if (NowTime.Hour == 21) HourWord = "час";
            else HourWord = "часа";

            if (NowTime.Minute == 1) MinetsWord = "минута";
            else if (NowTime.Minute > 1 && NowTime.Minute < 5) MinetsWord = "минуты";
            else if (NowTime.Minute > 4 && NowTime.Minute < 21) MinetsWord = "минут";
            else if (NowTime.Minute == 21) MinetsWord = "минута";
            else if (NowTime.Minute > 21 && NowTime.Minute < 25) MinetsWord = "минуты";
            else if (NowTime.Minute > 24 && NowTime.Minute < 31) MinetsWord = "минут";
            else if (NowTime.Minute == 31) MinetsWord = "минута";
            else if (NowTime.Minute > 31 && NowTime.Minute < 34) MinetsWord = "минуты";
            else if (NowTime.Minute > 34 && NowTime.Minute < 31) MinetsWord = "минут";
            else if (NowTime.Minute == 41) MinetsWord = "минута";
            else if (NowTime.Minute > 41 && NowTime.Minute < 45) MinetsWord = "минуты";
            else if (NowTime.Minute > 44 && NowTime.Minute < 51) MinetsWord = "минут";
            else if (NowTime.Minute == 51) MinetsWord = "минута";
            else if (NowTime.Minute > 51 && NowTime.Minute < 55) MinetsWord = "минуты";
            else MinetsWord = "минут";
            synth.Speak($"{NowTime.Hour} {HourWord}, {NowTime.Minute} {MinetsWord}  {str}");

        }
        static void ScreenShot()
        {
            try
            {
                Bitmap prt = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                Graphics graphics = Graphics.FromImage(prt as Image);
                graphics.CopyFromScreen(0, 0, 0, 0, prt.Size);
                Random rnd = new Random();
                int sk = rnd.Next(10000);
                prt.Save(@"D:\Скриншоты\" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            catch(FileNotFoundException ex)
            {
                Console.WriteLine(ex);
            }            
        }

        static void Voice(string word)
        {
            synth.SpeakAsyncCancelAll();
            synth.Speak(String.Format(word));
            //m_SoundPlayer.Play();
        }
        static void ProcStart(string dir, string word)
        {
            Console.Beep(1500, 200);
            try
            {      
                Voice(String.Format("Включаю " + word));
                //Thread.Sleep(3000);
                Process.Start(dir);                
            }
            catch(Exception e)
            {
                RecordEntry(e.Message);
            }            
            RecordEntry(word);
        }
        static void PlayMario()
        {
            Console.Beep(659, 125); Console.Beep(659, 125); Thread.Sleep(125);
            Console.Beep(659, 125); Thread.Sleep(167); Console.Beep(523, 125);
            Console.Beep(659, 125); Thread.Sleep(125); Console.Beep(784, 125);
            Thread.Sleep(375); Console.Beep(392, 125); Thread.Sleep(375);
            Console.Beep(523, 125); Thread.Sleep(250); Console.Beep(392, 125);
            Thread.Sleep(250); Console.Beep(330, 125); Thread.Sleep(250);
            Console.Beep(440, 125); Thread.Sleep(125); Console.Beep(494, 125);
            Thread.Sleep(125); Console.Beep(466, 125); Thread.Sleep(42);
            Console.Beep(440, 125); Thread.Sleep(125); Console.Beep(392, 125);
            Thread.Sleep(125); Console.Beep(659, 125); Thread.Sleep(125);
            Console.Beep(784, 125); Thread.Sleep(125); Console.Beep(880, 125);
            Thread.Sleep(125); Console.Beep(698, 125); Console.Beep(784, 125);
            Thread.Sleep(125); Console.Beep(659, 125); Thread.Sleep(125);
            Console.Beep(523, 125); Thread.Sleep(125); Console.Beep(587, 125);
            Console.Beep(494, 125); Thread.Sleep(125); Console.Beep(523, 125);
            Thread.Sleep(250); Console.Beep(392, 125); Thread.Sleep(250);
            Console.Beep(330, 125); Thread.Sleep(250); Console.Beep(440, 125);
            Thread.Sleep(125); Console.Beep(494, 125); Thread.Sleep(125);
            Console.Beep(466, 125); Thread.Sleep(42); Console.Beep(440, 125);
            Thread.Sleep(125); Console.Beep(392, 125); Thread.Sleep(125);
            Console.Beep(659, 125); Thread.Sleep(125); Console.Beep(784, 125);
            Thread.Sleep(125); Console.Beep(880, 125); Thread.Sleep(125);
            Console.Beep(698, 125); Console.Beep(784, 125); Thread.Sleep(125);
            Console.Beep(659, 125); Thread.Sleep(125); Console.Beep(523, 125);
            Thread.Sleep(125); Console.Beep(587, 125); Console.Beep(494, 125);
            Thread.Sleep(375); Console.Beep(784, 125); Console.Beep(740, 125);
            Console.Beep(698, 125); Thread.Sleep(42); Console.Beep(622, 125);
            Thread.Sleep(125); Console.Beep(659, 125); Thread.Sleep(167);
            Console.Beep(415, 125); Console.Beep(440, 125); Console.Beep(523, 125);
            Thread.Sleep(125); Console.Beep(440, 125); Console.Beep(523, 125);
            Console.Beep(587, 125); Thread.Sleep(250); Console.Beep(784, 125);
            Console.Beep(740, 125); Console.Beep(698, 125); Thread.Sleep(42);
            Console.Beep(622, 125); Thread.Sleep(125); Console.Beep(659, 125);
            Thread.Sleep(167); Console.Beep(698, 125); Thread.Sleep(125);
            Console.Beep(698, 125); Console.Beep(698, 125); Thread.Sleep(625);
            Console.Beep(784, 125); Console.Beep(740, 125); Console.Beep(698, 125);
            Thread.Sleep(42); Console.Beep(622, 125); Thread.Sleep(125);
            Console.Beep(659, 125); Thread.Sleep(167); Console.Beep(415, 125);
            Console.Beep(440, 125); Console.Beep(523, 125); Thread.Sleep(125);
            Console.Beep(440, 125); Console.Beep(523, 125); Console.Beep(587, 125);
            Thread.Sleep(250); Console.Beep(622, 125); Thread.Sleep(250);
            Console.Beep(587, 125); Thread.Sleep(250); Console.Beep(523, 125);
            Thread.Sleep(1125); Console.Beep(784, 125); Console.Beep(740, 125);
            Console.Beep(698, 125); Thread.Sleep(42); Console.Beep(622, 125);
            Thread.Sleep(125); Console.Beep(659, 125); Thread.Sleep(167);
            Console.Beep(415, 125); Console.Beep(440, 125); Console.Beep(523, 125);
            Thread.Sleep(125); Console.Beep(440, 125); Console.Beep(523, 125);
            Console.Beep(587, 125); Thread.Sleep(250); Console.Beep(784, 125);
            Console.Beep(740, 125); Console.Beep(698, 125); Thread.Sleep(42);
            Console.Beep(622, 125); Thread.Sleep(125); Console.Beep(659, 125);
            Thread.Sleep(167); Console.Beep(698, 125); Thread.Sleep(125);
            Console.Beep(698, 125); Console.Beep(698, 125); Thread.Sleep(625);
            Console.Beep(784, 125); Console.Beep(740, 125); Console.Beep(698, 125);
            Thread.Sleep(42); Console.Beep(622, 125); Thread.Sleep(125);
            Console.Beep(659, 125); Thread.Sleep(167); Console.Beep(415, 125);
            Console.Beep(440, 125); Console.Beep(523, 125); Thread.Sleep(125);
            Console.Beep(440, 125); Console.Beep(523, 125); Console.Beep(587, 125);
            Thread.Sleep(250); Console.Beep(622, 125); Thread.Sleep(250);
            Console.Beep(587, 125); Thread.Sleep(250); Console.Beep(523, 125);

        }
        static private void RecordEntry(string fileEvent)
        {
            //код внутри данного блока блокируется и становится недоступным для других потоков до завершения работы текущего потока
            lock (obj)
            {
                SW.WriteLine(String.Format("{0} {1}", DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"), fileEvent));
                SW.Flush();         //Очищает все буферы для текущего средства записи и вызывает запись всех данных буфера в основной поток
            }
        }
        static void HistoryReader(string Buftext)
        {
            if(Buftext != null)
            {
                synth.Speak(Buftext);
            }
            else
            {
                List<string> Text = new List<string>();

                string line = null;

                while ((line = SR.ReadLine()) != null)
                {
                    if (line != "")
                    {
                        Text.Add(line);
                    }
                }

                foreach (string str in Text)
                {
                    synth.Speak(str);
                }
            }
            
        }
    }
}
