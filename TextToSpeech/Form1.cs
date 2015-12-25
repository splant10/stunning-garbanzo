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
            radioButton1.Checked = true;
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
                synth.SpeakAsync(builder);
            }
        }

        // Choose Male voice
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                synth.SelectVoiceByHints(VoiceGender.Male);
            }
        }

        // Choose Female voice
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                synth.SelectVoiceByHints(VoiceGender.Female);
            }
        }

        // ///////////////// Cut, Copy, Paste, Undo \\\\\\\\\\\\\\\\\\\\\\ //
        // https://msdn.microsoft.com/en-us/library/system.windows.forms.textboxbase.paste(v=vs.110).aspx
        //
        // Undo
        private void undo()
        {
            // Determine if last operation can be undone in text box.   
            if (richTextBox1.CanUndo == true)
            {
                // Undo the last operation.
                richTextBox1.Undo();
                // Clear the undo buffer to prevent last action from being redone.
                richTextBox1.ClearUndo();
            }
        }
        //
        // Copy
        private void copy()
        {
            // Ensure that text is selected in the text box.   
            if (richTextBox1.SelectionLength > 0)
                // Copy the selected text to the Clipboard.
                richTextBox1.Copy();
        }
        //
        // Cut
        private void cut()
        {
            // Ensure that text is currently selected in the text box.   
            if (richTextBox1.SelectedText != "")
                // Cut the selected text in the control and paste it into the Clipboard.
                richTextBox1.Cut();
        }
        //
        // Paste
        private void paste()
        {
            // Determine if there is any text in the Clipboard to paste into the text box.
            if (Clipboard.GetDataObject().GetDataPresent(DataFormats.Text) == true)
            {
                // Determine if any text is selected in the text box.
                if (richTextBox1.SelectionLength > 0)
                {
                    // Move selection to the point after the current selection and paste.
                    richTextBox1.SelectionStart = richTextBox1.SelectionStart + richTextBox1.SelectionLength;
                }
                // Paste current text in Clipboard into text box.
                richTextBox1.Paste();
            }
        }
        // \\\\\\\\\\\\\\\\\\\\ End Cut, Copy, Paste, Undo ///////////////////// //



        // ////////////////////// Menu bar Edit \\\\\\\\\\\\\\\\\\\\\\\\\
        // 
        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            undo();
        }
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copy();
        }
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cut();
        }
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            paste();
        }
        // \\\\\\\\\\\\\\\\\\\ End  Menu bar Edit //////////////////////////



        // //////////////////// Right click menu \\\\\\\\\\\\\\\\\\\\\\\\
        private void undoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            undo();
        }
        private void copyToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            copy();
        }

        private void cutToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cut();
        }

        private void pasteToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            paste();
        }
        // \\\\\\\\\\\\\\\\\\\\\\ End Right click menu //////////////////
    }
}



// Pdf Text Getting
// from Parsing PDF Files using iTextSharp (C#, .NET):
// http://www.squarepdf.net/parsing-pdf-files-using-itextsharp
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
