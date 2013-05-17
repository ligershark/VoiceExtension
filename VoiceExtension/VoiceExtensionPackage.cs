using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;

namespace MadsKristensen.VoiceExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVoiceExtensionPkgString)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class VoiceExtensionPackage : Package
    {
        private DTE2 _dte;
        private SpeechRecognitionEngine _rec;
        private Dictionary<string, string> _dic;

        protected override void Initialize()
        {
            base.Initialize();
            _dte = GetService(typeof(DTE)) as DTE2;
            _dic = new Dictionary<string, string>();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID cmd = new CommandID(GuidList.guidVoiceExtensionCmdSet, (int)PkgCmdIDList.cmdidMyCommand);
                MenuCommand menu = new MenuCommand(OnListening, cmd);
                mcs.AddCommand(menu);
            }

            BuildCommandList();
            InitializeSpeechRecognition();
        }

        private void OnListening(object sender, EventArgs e)
        {
            _rec.RecognizeAsync();
            _dte.StatusBar.Text = "I'm listening...";
            _dte.StatusBar.Highlight(true);
        }

        private void InitializeSpeechRecognition()
        {
            var c = new Choices(_dic.Keys.ToArray());
            var gb = new GrammarBuilder(c);
            var g = new Grammar(gb);

            _rec = new SpeechRecognitionEngine(new CultureInfo("en"));
            _rec.RecognizeCompleted += OnRecognizeCompleted;
            _rec.InitialSilenceTimeout = TimeSpan.FromSeconds(3);

            _rec.LoadGrammar(g);
            _rec.SetInputToDefaultAudioDevice();
        }

        private void OnRecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            _rec.RecognizeAsyncStop();

            if (e.Result != null && e.Result.Confidence > 0.87 && _dic.ContainsKey(e.Result.Text))
            {
                if (TryExecuteCommand(_dic[e.Result.Text]))
                    _dte.StatusBar.Text = e.Result.Text;
                else
                    _dte.StatusBar.Text = e.Result.Text + " is not available in this context";
            }
            else if (e.Result != null)
            {
                _dte.StatusBar.Text = "Please repeat. I didn't quite understand";
            }
            else
            {
                _dte.StatusBar.Clear();
            }
        }

        private void BuildCommandList()
        {
            foreach (Command cmd in _dte.Commands)
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

                if (!string.IsNullOrEmpty(final) && !_dic.ContainsKey(final))
                    _dic.Add(final, cmd.Name);
            }
        }

        private bool TryExecuteCommand(string commandName)
        {
            var command = _dte.Commands.Item(commandName);

            if (command.IsAvailable)
                _dte.ExecuteCommand(command.Name);

            return command.IsAvailable;
        }
    }
}