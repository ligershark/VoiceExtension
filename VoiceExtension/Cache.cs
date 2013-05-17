using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;

namespace MadsKristensen.VoiceExtension
{
    internal class Cache
    {
        public static Dictionary<string, string> Commands { get; set; }

        public static void BuildCommandList(DTE2 dte)
        {
            Commands = new Dictionary<string, string>();

            foreach (Command cmd in dte.Commands)
            {
                string name = cmd.Name.Substring(cmd.Name.LastIndexOf('.') + 1);
                string final = string.Empty;

                for (int i = name.Length - 1; i > -1; i--)
                {
                    final = final.Insert(0, name[i].ToString());

                    if (i > 0 && char.IsUpper(name[i]))
                    {
                        final = final.Insert(0, " ");
                    }
                }

                if (!string.IsNullOrEmpty(final) && !Commands.ContainsKey(final))
                    Commands.Add(final, cmd.Name);
            }
        }
    }
}
