﻿using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;

namespace MadsKristensen.VoiceExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVoiceExtensionPkgString)]
    public sealed class VoiceExtensionPackage : Package
    {
        private const float _confidence = 0.8F;

        private DTE2 _dte;
        private SpeechRecognitionEngine _rec;

        protected override void Initialize()
        {
            base.Initialize();
            _dte = GetService(typeof(DTE)) as DTE2;

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                CommandID cmd = new CommandID(GuidList.guidVoiceExtensionCmdSet, (int)PkgCmdIDList.cmdidMyCommand);
                MenuCommand menu = new MenuCommand(OnListening, cmd);
                mcs.AddCommand(menu);
            }

            Cache.BuildCommandList(_dte);
            InitializeSpeechRecognition();
        }

        private void OnListening(object sender, EventArgs e)
        {
            _rec.RecognizeAsyncStop();
            _rec.RecognizeAsync();

            _dte.StatusBar.Text = "I'm listening...";
            _dte.StatusBar.Highlight(true);
        }

        private void InitializeSpeechRecognition()
        {
            var c = new Choices(Cache.Commands.Keys.ToArray());
            var gb = new GrammarBuilder(c);
            var g = new Grammar(gb);

            _rec = new SpeechRecognitionEngine();
            _rec.RecognizeCompleted += OnRecognizeCompleted;
            _rec.InitialSilenceTimeout = TimeSpan.FromSeconds(3);

            _rec.LoadGrammar(g);
            _rec.SetInputToDefaultAudioDevice();
        }

        private void OnRecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            _rec.RecognizeAsyncStop();

            if (e.Result != null && e.Result.Confidence > _confidence && Cache.Commands.ContainsKey(e.Result.Text))
            {
                if (!TryExecuteCommand(Cache.Commands[e.Result.Text], e.Result.Text))
                    _dte.StatusBar.Text = e.Result.Text + " is not available in this context";
            }
            else
            {
                _dte.StatusBar.Text = "I didn't quite get that. Please try again";
            }
        }

        private bool TryExecuteCommand(string commandName, string displayName)
        {
            var command = _dte.Commands.Item(commandName);

            if (command.IsAvailable)
            {
                _dte.ExecuteCommand(command.Name);

                int index = -1;
                var bindings = ((object[])command.Bindings).LastOrDefault();
                string keys = string.Empty;

                if (bindings != null)
                {
                    keys = bindings.ToString();
                    index = keys.IndexOf(':') + 2;
                }

                if (index > 1)
                    _dte.StatusBar.Text = displayName + " (" + keys.Substring(index) + ")";
                else
                    _dte.StatusBar.Text = displayName;
            }

            return command.IsAvailable;
        }
    }
}