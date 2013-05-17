using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace MadsKristensen.VoiceExtension
{
    internal class Cache
    {
        public static Dictionary<string, string> Commands { get; set; }

        public static void BuildCommandList(DTE2 dte)
        {
            Commands = new Dictionary<string, string>();

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string file = Path.Combine(folder, "resources", "commands.txt");

            if (File.Exists(file))
            {
                using (StreamReader reader = new StreamReader(file))
                {
                    string[] lines = reader.ReadToEnd().Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string line in lines)
                    {
                        if (line.StartsWith("#"))
                            continue;

                        string[] args = line.Split('|');

                        if (args.Length == 2)
                            AddCommand(args[1], args[0]);
                    }
                }
            }

            //foreach (Command cmd in dte.Commands)
            //{
            //    string name = cmd.Name.Substring(cmd.Name.LastIndexOf('.') + 1);
            //    string realName = string.Empty;

            //    for (int i = name.Length - 1; i > -1; i--)
            //    {
            //        realName = realName.Insert(0, name[i].ToString());

            //        if (i > 0 && char.IsUpper(name[i]))
            //        {
            //            realName = realName.Insert(0, " ");
            //        }
            //    }

            //    //AddCommand(cmd, realName);
            //}

            //using (var fr = new System.IO.StreamWriter("c:\\users\\madsk\\commands.txt"))
            //{
            //    foreach (string key in Commands.Keys)
            //    {
            //        fr.WriteLine(key + "|" + Commands[key]);
            //    }
            //}
        }

        private static void AddCommand(string commandName, string realName)
        {
            string clean = realName
                            .Replace("\"", string.Empty)
                            .Replace("'", string.Empty);

            if (!string.IsNullOrEmpty(clean) && !Commands.ContainsKey(clean))
            {
                Commands.Add(clean, commandName);
            }
        }
    }
}
