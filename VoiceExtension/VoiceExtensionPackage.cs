using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Windows.Forms;

namespace MadsKristensen.VoiceExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidVoiceExtensionPkgString)]
    public sealed class VoiceExtensionPackage : Package
    {
        private const float _minConfidence = 0.80F; // A value between 0 and 1

        private DTE2 _dte;
        private CommandTable _cache;
        private SpeechRecognitionEngine _rec;
        private bool _isEnabled, _isListening;

        protected override void Initialize()
        {
            _dte = GetService(typeof(DTE)) as DTE2;
            _cache = new CommandTable(_dte);
            InitializeSpeechRecognition();

            // Setup listening command
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            CommandID cmd = new CommandID(GuidList.guidVoiceExtensionCmdSet, (int)PkgCmdIDList.cmdidMyCommand);
            MenuCommand menu = new MenuCommand(OnListening, cmd);
            mcs.AddCommand(menu);
        }

        private void InitializeSpeechRecognition()
        {
            try
            {
                var c = new Choices(_cache.Commands.Keys.ToArray());
                var gb = new GrammarBuilder(c);
                var g = new Grammar(gb);

                _rec = new SpeechRecognitionEngine();
                _rec.InitialSilenceTimeout = TimeSpan.FromSeconds(3);
                _rec.SpeechHypothesized += OnSpeechHypothesized;
                _rec.RecognizeCompleted += OnSpeechRecognized;

                _rec.LoadGrammarAsync(g);
                _rec.SetInputToDefaultAudioDevice();

                _isEnabled = true;
            }
            catch { /* Speech Recognition hasn't been enabled on Windows */ }
        }

        private void OnListening(object sender, EventArgs e)
        {
            if (!_isEnabled)
            {
                SetupVoiceRecognition();
            }
            else if (!_isListening)
            {
                _isListening = true;
                _rec.RecognizeAsync();
                _dte.StatusBar.Text = "I'm listening...";
            }
        }

        private void OnSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            _dte.StatusBar.Text = "I'm listening... (" + e.Result.Text + " " + Math.Round(e.Result.Confidence * 100) + "%)";
        }

        private void OnSpeechRecognized(object sender, RecognizeCompletedEventArgs e)
        {
            _rec.RecognizeAsyncStop();
            _isListening = false;

            if (e.Result != null && e.Result.Confidence > _minConfidence)
            {
                _cache.ExecuteCommand(e.Result.Text);
            }
            else
            {
                _dte.StatusBar.Text = "I didn't quite get that. Please try again.";
            }
        }

        private static void SetupVoiceRecognition()
        {
            string message = "Do you want to learn how to setup voice recognition in Windows?";

            if (MessageBox.Show(message, "Voice Commands for Visual Studio", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                using (var process = new System.Diagnostics.Process())
                {
                    process.StartInfo = new ProcessStartInfo("http://windows.microsoft.com/en-US/windows-8/using-speech-recognition/");
                    process.Start();
                }
            }
        }
    }
}