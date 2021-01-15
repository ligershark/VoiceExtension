using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using EnvDTE;
using EnvDTE80;

namespace MadsKristensen.VoiceExtension
{
    internal class CommandTable
    {
        private readonly DTE2 _dte;

        public CommandTable(DTE2 dte)
        {
            _dte = dte;
            BuildCommandTable();
        }

        public Dictionary<string, string> Commands { get; private set; }

        private void BuildCommandTable()
        {
            Commands = new Dictionary<string, string> { { "yes", null }, { "no", null }, { "what can I say", null } };

            var folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var file = Path.Combine(folder, "resources", "commands.txt");

            if (File.Exists(file))
            {
                var lines = File.ReadAllLines(file);

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var args = line.Split('|');

                    if (args.Length == 2)
                    {
                        AddCommand(args[1], args[0]);
                    }
                }
            }
        }

        //private void GenerateCommandFile()
        //{
        //    foreach (Command cmd in _dte.Commands)
        //    {
        //        string name = cmd.Name.Substring(cmd.Name.LastIndexOf('.') + 1);
        //        string realName = string.Empty;

        //        for (int i = name.Length - 1; i > -1; i--)
        //        {
        //            realName = realName.Insert(0, name[i].ToString());

        //            if (i > 0 && char.IsUpper(name[i]))
        //            {
        //                realName = realName.Insert(0, " ");
        //            }
        //        }

        //        AddCommand(cmd.Name, realName);
        //    }

        //    using (var fr = new System.IO.StreamWriter("c:\\users\\madsk\\commands.txt"))
        //    {
        //        foreach (string key in Commands.Keys)
        //        {
        //            fr.WriteLine(key + "|" + Commands[key].Name);
        //        }
        //    }
        //}

        private void AddCommand(string commandName, string realName)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var clean = realName
                            .Replace("\"", string.Empty)
                            .Replace("'", string.Empty);

            if (!string.IsNullOrEmpty(clean) && !Commands.ContainsKey(clean))
            {
                try
                {
                    Command command = _dte.Commands.Item(commandName);
                    Commands.Add(clean, commandName);
                }
                catch (Exception ex)
                {
                    Trace.Write(ex.ToString());
                }
            }
        }

        public void ExecuteCommand(string displayName)
        {
            if (!Commands.ContainsKey(displayName))
            {
                return;
            }

            var realName = Commands[displayName];
            Command command = _dte.Commands.Item(realName);

            if (command != null && command.IsAvailable)
            {
                _dte.ExecuteCommand(command.Name);
                DisplayKeyBindings(displayName, command);
            }
            else
            {
                _dte.StatusBar.Text = displayName + " is not available in this context";
            }
        }

        private void DisplayKeyBindings(string displayName, Command command)
        {
            _dte.StatusBar.Text = displayName;
            var bindings = ((object[])command.Bindings).FirstOrDefault() as string;

            if (!string.IsNullOrEmpty(bindings))
            {
                var index = bindings.IndexOf(':') + 2;
                _dte.StatusBar.Text += " (" + bindings.Substring(index) + ")";
            }
        }
    }
}
