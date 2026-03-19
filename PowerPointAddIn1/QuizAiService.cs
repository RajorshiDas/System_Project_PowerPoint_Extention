using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace PowerPointAddIn1
{
    public class QuizAiService
    {
        private readonly string _apiKey;
        private const string ModelName = "gemini-2.0-flash";

        public QuizAiService(string apiKey)
        {
            _apiKey = apiKey;
        }

        public async Task<QuizSet> GenerateQuizAsync(string slideText, int questionCount)
        {
            string prompt =
                "You are a quiz generator.\n" +
                "Generate quiz questions ONLY from the slide text below.\n" +
                "Return ONLY valid JSON.\n\n" +
                "Required JSON format:\n" +
                "{\n" +
                "  \"Title\": \"Quiz\",\n" +
                "  \"Questions\": [\n" +
                "    {\n" +
                "      \"Question\": \"...\",\n" +
                "      \"Options\": [\"...\", \"...\", \"...\", \"...\"],\n" +
                "      \"CorrectAnswer\": \"...\",\n" +
                "      \"Explanation\": \"...\"\n" +
                "    }\n" +
                "  ]\n" +
                "}\n\n" +
                "Rules:\n" +
                "- Create exactly " + questionCount + " multiple-choice questions\n" +
                "- Each question must have exactly 4 options\n" +
                "- CorrectAnswer must exactly match one option\n" +
                "- Do not include markdown\n" +
                "- Do not include any text outside JSON\n\n" +
                "Slide text:\n" + slideText;

            string url =
                "https://generativelanguage.googleapis.com/v1beta/models/" +
                ModelName +
                ":generateContent?key=" + _apiKey;

            string body = JsonSerializer.Serialize(new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new[]
                        {
                            new { text = prompt }
                        }
                    }
                }
            });

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.SendAsync(request);
                string responseText = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception("Gemini API error:\n" + responseText);
                }

                string modelText = ExtractModelText(responseText);
                string json = ExtractJsonOnly(modelText);

                if (string.IsNullOrWhiteSpace(json))
                    return null;

                try
                {
                    return JsonSerializer.Deserialize<QuizSet>(json, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });
                }
                catch
                {
                    return null;
                }
            }
        }

        private string ExtractModelText(string responseText)
        {
            using (JsonDocument doc = JsonDocument.Parse(responseText))
            {
                JsonElement root = doc.RootElement;

                if (!root.TryGetProperty("candidates", out JsonElement candidates))
                    return "";

                if (candidates.GetArrayLength() == 0)
                    return "";

                JsonElement firstCandidate = candidates[0];

                if (!firstCandidate.TryGetProperty("content", out JsonElement content))
                    return "";

                if (!content.TryGetProperty("parts", out JsonElement parts))
                    return "";

                if (parts.GetArrayLength() == 0)
                    return "";

                JsonElement firstPart = parts[0];

                if (!firstPart.TryGetProperty("text", out JsonElement textElement))
                    return "";

                return textElement.GetString() ?? "";
            }
        }

        private string ExtractJsonOnly(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return "";

            int start = raw.IndexOf("{");
            int end = raw.LastIndexOf("}");

            if (start >= 0 && end > start)
                return raw.Substring(start, end - start + 1);

            return "";
        }
    }
}