using System.Collections.Generic;

namespace PowerPointAddIn1
{
    public class QuizSet
    {
        public string Title { get; set; }
        public List<QuizQuestion> Questions { get; set; }

        public QuizSet()
        {
            Title = "Quiz";
            Questions = new List<QuizQuestion>();
        }
    }
}