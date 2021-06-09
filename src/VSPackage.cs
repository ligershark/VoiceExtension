using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Threading;
using System.Windows;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace MadsKristensen.VoiceExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidVoiceExtensionPkgString)]
    public sealed class VSPackage : AsyncPackage
    {
        private const float _minConfidence = 0.80F; // A value between 0 and 1

        private DTE2 _dte;
        private CommandTable _cache;
        private SpeechRecognitionEngine _rec;
        private bool _isEnabled, _isListening;
        private string _rejected;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            _dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(_dte);

            _cache = new CommandTable(_dte);
            OutputWindowTraceListener.Register(Vsix.Name, nameof(VoiceExtension));

            InitializeSpeechRecognition();

            // Setup listening command
            var mcs = await GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Assumes.Present(mcs);

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

                _rec = new SpeechRecognitionEngine
                {
                    InitialSilenceTimeout = TimeSpan.FromSeconds(3)
                };
                _rec.SpeechHypothesized += OnSpeechHypothesized;
                _rec.SpeechRecognitionRejected += OnSpeechRecognitionRejected;
                _rec.RecognizeCompleted += OnSpeechRecognized;

                _rec.LoadGrammarAsync(g);
                _rec.SetInputToDefaultAudioDevice();

                _isEnabled = true;
            }
            catch (Exception ex)
            {
                Trace.Write(ex.ToString());
            }
        }

        private void OnListening(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
                Trace.Write(ex.ToString());
            }
        }

        private void OnSpeechHypothesized(object sender, SpeechHypothesizedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrEmpty(_rejected))
            {
                _dte.StatusBar.Text = "I'm listening... (" + e.Result.Text + " " + Math.Round(e.Result.Confidence * 100) + "%)";
            }
        }

        private void OnSpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e.Result.Text != "yes" && e.Result.Confidence > 0.5F)
            {
                _rejected = e.Result.Text;
                _dte.StatusBar.Text = "Did you mean " + e.Result.Text + "? (say yes or no)";
            }
        }

        private void OnSpeechRecognized(object sender, RecognizeCompletedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                _rec.RecognizeAsyncStop();
                _isListening = false;

                if (e.Result != null && !string.IsNullOrEmpty(_rejected))
                { // Handle answer to precision question
                    _dte.StatusBar.Clear();

                    if (e.Result.Text == "yes")
                    {
                        _cache.ExecuteCommand(_rejected);
                    }

                    _rejected = null;
                }
                else if (e.Result != null && e.Result.Text == "what can I say")
                {// Show link to command list
                    System.Diagnostics.Process.Start("https://github.com/ligershark/VoiceExtension/blob/master/src/Resources/commands.txt");
                    _dte.StatusBar.Clear();
                }
                else if (e.Result != null && e.Result.Confidence > _minConfidence)
                { // Speech matches a command
                    _cache.ExecuteCommand(e.Result.Text);
                    var props = new Dictionary<string, string> { { "phrase", e.Result.Text } };
                }
                else if (string.IsNullOrEmpty(_rejected))
                { // Speech didn't match a command
                    _dte.StatusBar.Text = "I didn't quite get that. Please try again.";
                }
                else if (e.Result == null && !string.IsNullOrEmpty(_rejected) && !e.InitialSilenceTimeout)
                { // Keep listening when asked about rejected speech
                    _rec.RecognizeAsync();
                    //Telemetry.TrackEvent("Low confidence");
                }
                else
                { // No match or timeout
                    _dte.StatusBar.Clear();
                }
            }
            catch (Exception ex)
            {
                _dte.StatusBar.Clear();
                Trace.Write(ex.ToString());
            }
        }

        private static void SetupVoiceRecognition()
        {
            var message = "Do you want to learn how to setup voice recognition in Windows?";
            MessageBoxResult answer = MessageBox.Show(message, Vsix.Name, MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (answer == MessageBoxResult.Yes)
            {
                var url = "http://windows.microsoft.com/en-US/windows-8/using-speech-recognition/";
                System.Diagnostics.Process.Start(url);
            }
        }
    }
}