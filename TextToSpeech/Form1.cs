/*
 * Created by Spencer Plant Dec 13 2015
 * A text to speech app that allows a user to paste in text to save as a .wav file for later listening
 * Heavy use of SpeechSynthesizer, and this reference:
 * https://msdn.microsoft.com/en-us/library/ms586885(v=vs.110).aspx
 * 
 *    Copyright 2015 Spencer Plant
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
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
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
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

                string filetype = System.IO.Path.GetExtension(filename);
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
                // TODO Should change to a "pause" button but need to listen for SpeakCompleted event.

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

namespace TextToSpeech
{
    public static class PdfTextGetter
    {
        public static string getPdfText(string path)
        {
            PdfReader reader = new PdfReader(path);
            string text = string.Empty;
            for (int page = 1; page <= reader.NumberOfPages; page++)
            {
                text += PdfTextExtractor.GetTextFromPage(reader, page);
            }
            reader.Close();
            return text;
        }
    }
}
