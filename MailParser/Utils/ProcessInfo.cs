using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Logger;
using MailParser;

namespace System.Diagnostics
{
    public class ProcessInfo
    {
        public Process m_process { get; private set; }
        public int Id { get; private set; }
        public string ProcessName { get; private set; }

        public ProcessState ProcessState { get; private set; }

        private ProcessInfo(int pid)
        {
            Id = pid;
            m_process = Process.GetProcessById(pid);
            if (m_process == null)
            {
                throw new Exception("GetProcessById returned null");
            }
        }
        public static string get_current_process_name()
        {
            string process_name = Process.GetCurrentProcess().ProcessName;
            int pid = Process.GetCurrentProcess().Id;

            if (ConstEnv.OS_TYPE == ConstEnv.OS_WINDOWS)
                return process_name;

            var stat = GetStat(pid);
            if (stat == null)
                return process_name;

            // look for mono-specific processes
            if (stat.FileName == "(mono)" || stat.FileName == "(cli)")
            {
                var cmdline = GetCommandLine(stat.PID);
                foreach (string arg in cmdline)
                    MyLogger.Info($"mono process stat cmd line : {arg}");
                process_name = Path.GetFileName(cmdline[1]);
            }
            else
            {
                // trim out the parens
                process_name = stat.FileName.Trim(new char[] { '(', ')' });
            }
            return process_name;
        }
        public static ProcessInfo[] find_working_process()
        {
            ProcessInfo[] found_list = GetProcessesByName(get_current_process_name(), Process.GetCurrentProcess().Id);
            return found_list;
        }
        public static ProcessInfo[] GetProcessesByName(string processName, int except_pid = -1)
        {
            if (ConstEnv.OS_TYPE == ConstEnv.OS_WINDOWS)
                return GetProcessesForWin(except_pid).Where(p => p.ProcessName == processName).ToArray();
            else
                return GetProcessesForLinux(except_pid).Where(p => p.ProcessName == processName).ToArray();
        }
        public static ProcessInfo[] GetProcessesForWin(int except_pid = -1)
        {
            var list = new List<ProcessInfo>();
            foreach (Process process in Process.GetProcesses())
            {
                //MyLogger.Info($"enum process name = {process.ProcessName}, PID = {process.Id}");
                if (except_pid != process.Id)
                {
                    //if (process.StartInfo.Arguments == "-start work")
                    var proc = new ProcessInfo(process.Id);
                    proc.ProcessName = process.ProcessName;
                    list.Add(proc);
                }
            }
            return list.ToArray();
        }
        public static ProcessInfo[] GetProcessesForLinux(int except_pid = -1)
        {
            var list = new List<ProcessInfo>();

            foreach (var path in Directory.GetDirectories("/proc"))
            {
                var d = Path.GetFileName(path);
                int pid;

                if (!int.TryParse(d, out pid))
                {
                    continue;
                }
                if (pid == except_pid)
                    continue;

                // stat
                var stat = GetStat(pid);
                if (stat == null) continue;

                var proc = new ProcessInfo(stat.PID);
                proc.ProcessState = stat.State;

                //MyLogger.Info($"enum process stat : file name = {stat.FileName}, PID = {stat.PID}");

                // look for mono-specific processes
                if (stat.FileName == "(mono)" || stat.FileName == "(cli)")
                {
                    // TODO: handle command-line args to the Mono app
                    var cmdline = GetCommandLine(stat.PID);

                    MyLogger.Info($"mono process cmd : {cmdline.ToString()}");

                    // cmdline[0] == path to mono
                    // cmdline[1] == mono app
                    // cmdline[1+n] == mono app args
                    proc.ProcessName = Path.GetFileName(cmdline[1]);
                    MyLogger.Info($"mono process process name : {proc.ProcessName}");
                    for (int i = 2; i < cmdline.Length; i++)
                        MyLogger.Info($"{i - 1}th mono process args : {cmdline[i]}");
                }
                else
                {
                    // trim out the parens
                    proc.ProcessName = stat.FileName.Trim(new char[] { '(', ')' });
                    //MyLogger.Info($"non-mono process process name : {proc.ProcessName}");
                }

                list.Add(proc);
            }
            return list.ToArray();
        }

        private static Stat GetStat(int pid)
        {
            try
            {
                var statDir = string.Format("/proc/{0}/stat", pid);
                if (!File.Exists(statDir))
                    return null;

                var proc = new ProcessInfo(pid);

                using (var reader = File.OpenText(statDir))
                {
                    var line = reader.ReadToEnd();
                    return new Stat(line);
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
                return null;
            }
        }

        private static string[] GetCommandLine(int pid)
        {
            // The command line arguments appear in this file as a set of null-separated strings, with a further null byte after the last string. 
            using (var reader = File.OpenText(string.Format("/proc/{0}/cmdline", pid)))
            {
                string contents = reader.ReadToEnd();
                var args = contents.Split(new char[] { '\0' }, StringSplitOptions.RemoveEmptyEntries);
                return args;
            }
        }

    }

    internal class Stat
    {
        internal Stat(string procLine)
        {
            try
            {
                var items = procLine.Split(new char[] { ' ' }, StringSplitOptions.None);

                PID = int.Parse(items[0]);
                FileName = items[1];

                switch (items[2][0])
                {
                    case 'R':
                        State = ProcessState.Running;
                        break;
                    case 'S':
                        State = ProcessState.InterruptableWait;
                        break;
                    case 'D':
                        State = ProcessState.UninterruptableDiskWait;
                        break;
                    case 'Z':
                        State = ProcessState.Zombie;
                        break;
                    case 'T':
                        State = ProcessState.Traced;
                        break;
                    case 'W':
                        State = ProcessState.Paging;
                        break;
                }
            }
            catch (Exception exception)
            {
                MyLogger.Error($"Exception Error ({System.Reflection.MethodBase.GetCurrentMethod().Name}): {exception.Message + "\n" + exception.StackTrace}");
            }
        }

        public int PID { get; private set; }
        public string FileName { get; private set; }
        public ProcessState State { get; private set; }
    }

    public enum ProcessState
    {
        Running, // R
        InterruptableWait, // S
        UninterruptableDiskWait, // D
        Zombie, // Z
        Traced, // T
        Paging // W
    }
}