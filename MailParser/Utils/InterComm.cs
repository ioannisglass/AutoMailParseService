using MailParser;
using Logger;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Utils
{
    class InterComm
    {
        static private readonly string inter_comm_file_name = "intcomm.txt";

        static public readonly string inter_commmand_kill = "KILL";

        static public ProcessInfo[] find_working_process()
        {
            ProcessInfo[] found_list = ProcessInfo.find_working_process();
            MyLogger.Info($"Found working process count : {found_list.Length}");
            return found_list;
        }
        static private void write_to_intercomm_file(string data)
        {
            for (int i = 0; i < 300; i++)
            {
                try
                {
                    File.WriteAllText(inter_comm_file_name, data);
                    break;
                }
                catch
                {
                    Thread.Sleep(100);
                }
            }
        }
        static private string read_from_intercomm_file()
        {
            string data = "";

            for (int i = 0; i < 300; i++)
            {
                try
                {
                    data = File.ReadAllText(inter_comm_file_name);
                    break;
                }
                catch
                {
                    Thread.Sleep(100);
                }
            }
            return data;
        }
        static public void clear_intercomm_file()
        {
            File.WriteAllText(inter_comm_file_name, "");
        }
        static public void kill_working_Process()
        {
            ProcessInfo[] work_process = find_working_process();
            if (work_process.Length == 0)
                return;

            MyLogger.Info("Write [KILL] command.");
            write_to_intercomm_file(inter_commmand_kill);

            for (int i = 0; i < 600; i++)
            {
                Thread.Sleep(1000);
                if (find_working_process().Length == 0)
                {
                    MyLogger.Info("No worker process.");
                    break;
                }
            }

            // To Do. Force kill process.

            clear_intercomm_file();
        }
        static public void start_working_Process()
        {
            Process WorkerProcess = new Process();

            WorkerProcess.StartInfo.FileName = Assembly.GetExecutingAssembly().Location;
            WorkerProcess.StartInfo.Arguments = "-start work";
            WorkerProcess.StartInfo.UseShellExecute = false;
            WorkerProcess.StartInfo.CreateNoWindow = false;
            WorkerProcess.Start();

            MyLogger.Info("Start worker process.");
        }
        static public string get_interprocess_command()
        {
            string command = read_from_intercomm_file();
            return command;
        }
        static public void monitor_interprocess_command(CancellationTokenSource cts)
        {
            new Thread(() =>
            {
                MyLogger.Info("Monitoring interprocess command thread started.");

                while (!cts.IsCancellationRequested && !Program.g_must_end)
                {
                    string command = get_interprocess_command();
                    if (command != "")
                    {
                        MyLogger.Info($"Take command : {command}");
                    }

                    if (command == inter_commmand_kill)
                    {
                        MyLogger.Info($"Set termination flag as TRUE.");
                        cts.Cancel();
                        Program.g_must_end = true;
                        break;
                    }

                    Thread.Sleep(100);
                }

                MyLogger.Info("Monitoring interprocess command thread stopped.");

            }).Start();
        }
    }
}
