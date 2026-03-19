using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace PowerPointAddIn1
{
    public partial class QuizPaneControl : UserControl
    {
        public QuizPaneControl()
        {
            InitializeComponent();
        }

        private void QuizPaneControl_Load(object sender, EventArgs e)
        {

        }

        private void numQuestionCount_ValueChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                SlideTextService slideTextService = new SlideTextService();
                PowerPoint.Slide slide = slideTextService.GetActiveSlide(app);

                if (slide == null)
                {
                    MessageBox.Show("No active slide found.");
                    return;
                }

                string slideText = slideTextService.ExtractTextFromSlide(slide);

                if (string.IsNullOrWhiteSpace(slideText))
                {
                    MessageBox.Show("No text found on the current slide.");
                    return;
                }
                string apiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    MessageBox.Show("GEMINI_API_KEY is not set.");
                    return;
                }

                SetGenerateButtonEnabled(false);
                SetOutputText("Generating quiz...");

                int questionCount = (int)numQuestionCount.Value;

                QuizAiService quizAiService = new QuizAiService(apiKey);
                QuizSet quiz = await quizAiService.GenerateQuizAsync(slideText, questionCount);

                if (quiz == null || quiz.Questions == null || quiz.Questions.Count == 0)
                {
                    SetOutputText("Quiz generation failed.");
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(quiz.Title);
                sb.AppendLine();

                for (int i = 0; i < quiz.Questions.Count; i++)
                {
                    QuizQuestion q = quiz.Questions[i];

                    sb.AppendLine((i + 1) + ". " + q.Question);

                    for (int j = 0; j < q.Options.Count; j++)
                    {
                        char label = (char)('A' + j);
                        sb.AppendLine("   " + label + ". " + q.Options[j]);
                    }

                    sb.AppendLine("Answer: " + q.CorrectAnswer);
                    sb.AppendLine("Explanation: " + q.Explanation);
                    sb.AppendLine();
                }

                SetOutputText(sb.ToString());
            }
            catch (Exception ex)
            {
                string msg = ex.ToString();

                if (msg.Contains("RESOURCE_EXHAUSTED") || msg.Contains("quota") || msg.Contains("429"))
                {
                    MessageBox.Show(
                        "Gemini quota/rate limit reached. Try again later, reduce question count, or enable billing in Google AI Studio.",
                        "Quota Reached");
                    SetOutputText("Quota reached. Try again later.");
                }
                else
                {
                    MessageBox.Show(ex.Message, "Error");
                    SetOutputText("Error generating quiz.");
                }
            }
            finally
            {
                SetGenerateButtonEnabled(true);
            }
        }


        private void rtbQuizOutput_TextChanged(object sender, EventArgs e)
        {

        }
        private void SetOutputText(string text)
        {
            if (rtbQuizOutput.InvokeRequired)
            {
                rtbQuizOutput.Invoke(new Action(() => rtbQuizOutput.Text = text));
            }
            else
            {
                rtbQuizOutput.Text = text;
            }
        }

        private void SetGenerateButtonEnabled(bool enabled)
        {
            if (button1.InvokeRequired)
            {
                button1.Invoke(new Action(() => button1.Enabled = enabled));
            }
            else
            {
                button1.Enabled = enabled;
            }
        }
    }
}
