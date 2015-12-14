/*
 * Created by Spencer Plant Dec 13 2015
 * A text to speech app that allows a user to paste in text to save as a .wav file for later listening
 * Heavy use of SpeechSynthesizer, and this reference:
 * https://msdn.microsoft.com/en-us/library/ms586885(v=vs.110).aspx
 */


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

using System.IO;
using System.Speech.Synthesis;
using System.Speech.AudioFormat;

namespace TextToSpeech
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Load file button
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            // Set Filter options for dialog box
            openFileDialog1.Filter = "PDF|*.pdf|Text|*.txt|Rich Text Format|*.rtf";
            openFileDialog1.FilterIndex = 1;

            // Show load file dialog
            DialogResult result = openFileDialog1.ShowDialog();
            // Test the result
            if (result == DialogResult.OK)
            {
                // Get filename and location
                string filename = openFileDialog1.FileName;
                fileNameBox.Text = filename;

                string filetype = Path.GetExtension(filename);
                Console.WriteLine(filetype);
                // Handle the nasty little case of PDFs
                if (filetype == ".pdf")
                {
                    try
                    {
                        string pdfText = PdfTextGetter.getPdfText(filename);
                        richTextBox1.Text = pdfText;
                    } catch {
                        richTextBox1.Text = "There was a problem loading that file.";
                    }
                }
                else // RTF or Text format
                {
                    string filetext = System.IO.File.ReadAllText(@filename);
                    richTextBox1.Text = filetext;
                }
            }
        }

        static string tempfileLocation = System.AppDomain.CurrentDomain.BaseDirectory.ToString() + "\\temp.wav";
        static SpeechSynthesizer synth = new SpeechSynthesizer();

        // Play button click
        // Doesn't yet account for blank (spaces) input
        int playOrPause = 0; // 0 = Play button is displayed, 1 = Pause button is displayed
        private void button2_Click(object sender, EventArgs e)
        {
            if (playOrPause == 0) // Play button is clicked. 
            {
                // Should change to a "pause" button but need to listen for SpeakCompleted event.

                if (richTextBox1.Text.Length > 0)
                {
                    // Configure the audio output. 
                    synth.SetOutputToDefaultAudioDevice();

                    // Build a prompt. This is so the speaker doesn't pause after EVERY line
                    PromptBuilder builder = new PromptBuilder();
                    builder.AppendText(richTextBox1.Text);

                    // Speak the prompt
                    synth.SpeakAsync(builder);
                }
            }
        }

        // Stop button click
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                synth.SpeakAsyncCancelAll();
            }
            catch
            {
            }
        }

        // Save button click
        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = ".WAV|*.wav";
            saveFileDialog1.RestoreDirectory = true;
            DialogResult result = saveFileDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                string filename = saveFileDialog1.FileName;
                synth.SetOutputToWaveFile(filename, new SpeechAudioFormatInfo(32000, AudioBitsPerSample.Sixteen, AudioChannel.Mono));

                // Build a prompt. This is so the speaker doesn't pause after EVERY line
                PromptBuilder builder = new PromptBuilder();
                builder.AppendText(richTextBox1.Text);

                // Speak to the wav file
                synth.Speak(builder);
            }
        }
    }
}
