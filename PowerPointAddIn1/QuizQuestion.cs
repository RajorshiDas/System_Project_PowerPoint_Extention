using System.Collections.Generic;

namespace PowerPointAddIn1
{
    public class QuizQuestion
    {
        public string Question { get; set; }
        public List<string> Options { get; set; }
        public string CorrectAnswer { get; set; }
        public string Explanation { get; set; }

        public QuizQuestion()
        {
            Question = "";
            Options = new List<string>();
            CorrectAnswer = "";
            Explanation = "";
        }
    }
}