using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace MadsKristensen.VoiceExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidVoiceExtensionPkgString)]
    public sealed class VoiceExtensionPackage : Package
    {
        private const float _minConfidence = 0.80F; // A value between 0 and 1

        private DTE2 _dte;
        private CommandTable _cache;
        private SpeechRecognitionEngine _rec;
        private bool _isEnabled, _isListening;
        private string _rejected;

        protected override void Initialize()
        {
            _dte = GetService(typeof(DTE)) as DTE2;
            _cache = new CommandTable(_dte);

            Logger.Initialize(this, Vsix.Name);
            Telemetry.Initialize(this, Vsix.Version, "c18e8661-f6e6-466d-b968-a6c128506bf4");

            InitializeSpeechRecognition();

            // Setup listening command
            var mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            var cmd = new CommandID(PackageGuids.guidVoiceExtensionCmdSet, PackageIds.cmdidMyCommand);
            var menu = new MenuCommand(OnListening, cmd);
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
                _rec.SpeechRecognitionRejected += OnSpeechRecognitionRejected;
                _rec.RecognizeCompleted += OnSpeechRecognized;

                _rec.LoadGrammarAsync(g);
                _rec.SetInputToDefaultAudioDevice();

                _isEnabled = true;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private void OnListening(object sender, EventArgs e)
        {
            try
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
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private void OnSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            if (string.IsNullOrEmpty(_rejected))
                _dte.StatusBar.Text = "I'm listening... (" + e.Result.Text + " " + Math.Round(e.Result.Confidence * 100) + "%)";
        }

        private void OnSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            if (e.Result.Text != "yes" && e.Result.Confidence > 0.5F)
            {
                _rejected = e.Result.Text;
                _dte.StatusBar.Text = "Did you mean " + e.Result.Text + "? (say yes or no)";
            }
        }

        private void OnSpeechRecognized(object sender, RecognizeCompletedEventArgs e)
        {
            try
            {
                _rec.RecognizeAsyncStop();
                _isListening = false;

                if (e.Result != null && !string.IsNullOrEmpty(_rejected))
                { // Handle answer to precision question
                    _dte.StatusBar.Clear();

                    if (e.Result.Text == "yes")
                        _cache.ExecuteCommand(_rejected);

                    _rejected = null;
                }
                else if (e.Result != null && e.Result.Text == "what can I say")
                {// Show link to command list
                    System.Diagnostics.Process.Start("https://raw.github.com/ligershark/VoiceExtension/master/VoiceExtension/Resources/commands.txt");
                    _dte.StatusBar.Clear();
                    Telemetry.TrackEvent("Show command list");
                }
                else if (e.Result != null && e.Result.Confidence > _minConfidence)
                { // Speech matches a command
                    _cache.ExecuteCommand(e.Result.Text);
                    var props = new Dictionary<string, string> { { "phrase" , e.Result.Text } };
                    Telemetry.TrackEvent("Match", props);
                }
                else if (string.IsNullOrEmpty(_rejected))
                { // Speech didn't match a command
                    _dte.StatusBar.Text = "I didn't quite get that. Please try again.";
                    Telemetry.TrackEvent("No match");
                }
                else if (e.Result == null && !string.IsNullOrEmpty(_rejected) && !e.InitialSilenceTimeout)
                { // Keep listening when asked about rejected speech
                    _rec.RecognizeAsync();
                    Telemetry.TrackEvent("Low confidence");
                }
                else
                { // No match or timeout
                    _dte.StatusBar.Clear();
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private static void SetupVoiceRecognition()
        {
            string message = "Do you want to learn how to setup voice recognition in Windows?";
            var answer = MessageBox.Show(message, Vsix.Name, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (answer == DialogResult.Yes)
            {
                string url = "http://windows.microsoft.com/en-US/windows-8/using-speech-recognition/";
                System.Diagnostics.Process.Start(url);
            }
        }
    }
}