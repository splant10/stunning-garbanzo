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

                }
            }
        }
    }
}
