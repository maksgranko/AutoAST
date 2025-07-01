using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace AutoAST
{
    public class Question
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public List<string> CorrectAnswers { get; set; }
        public List<string> WrongAnswers { get; set; }
        public bool IsOpenQuestion { get; set; }
        public string Topic { get; set; }
    }

    public class AstParser : IDisposable
    {
        private readonly string _dbPath;

        public void Dispose()
        {
            if (_connection != null)
            {
                _connection.Dispose();
                _connection = null;
            }
        }

        ~AstParser()
        {
            Dispose();
        }
        private readonly Dictionary<string, string> _connectionStrings;
        private string _activeProvider;
        private Dictionary<int, string> _topics;
        private OleDbConnection _connection;

        public AstParser(string dbPath)
        {
            _dbPath = dbPath;
            _connectionStrings = new Dictionary<string, string>
            {
                ["Microsoft.Jet.OLEDB.4.0"] = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={dbPath};Persist Security Info=False;",
                ["Microsoft.ACE.OLEDB.12.0"] = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;"
            };
        }

        private void EnsureConnectionOpen()
        {
            if (_connection != null && _connection.State == ConnectionState.Open)
            {
                return;
            }

            if (_connection != null)
            {
                _connection.Dispose();
                _connection = null;
            }

            if (_activeProvider != null)
            {
                _connection = new OleDbConnection(_connectionStrings[_activeProvider]);
                _connection.Open();
                return;
            }

            foreach (var provider in _connectionStrings.Keys)
            {
                try
                {
                    _connection = new OleDbConnection(_connectionStrings[provider]);
                    _connection.Open();
                    _activeProvider = provider;
                    Console.WriteLine($"Успешное подключение через провайдер {provider}");
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Не удалось подключиться через провайдер {provider}: {ex.Message}");
                    if (_connection != null)
                    {
                        _connection.Dispose();
                        _connection = null;
                    }
                }
            }

            throw new Exception("Не удалось подключиться к базе данных. Установите необходимые компоненты Access Database Engine.");
        }

        private OleDbConnection GetConnection()
        {
            EnsureConnectionOpen();
            return _connection;
        }

        private Dictionary<int, string> LoadTopics()
        {
            var topics = new Dictionary<int, string>();
            try
            {
                EnsureConnectionOpen();
                var command = new OleDbCommand("SELECT id1, name FROM Level1", _connection);
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var id = reader.GetInt32(0);
                        var name = reader.GetString(1);
                        topics[id] = name;
                    }
                }
                return topics;
            }
            catch (Exception)
            {
                if (_connection != null)
                {
                    _connection.Dispose();
                    _connection = null;
                }
                throw;
            }
        }

        public List<Question> ParseQuestions()
        {
            try
            {
                var questions = new List<Question>();
                _topics = LoadTopics();

                EnsureConnectionOpen();
                questions.AddRange(ParseOpenQuestions(_connection));
                questions.AddRange(ParseClosedQuestions(_connection));

                return questions;
            }
            catch (Exception)
            {
                if (_connection != null)
                {
                    _connection.Dispose();
                    _connection = null;
                }
                throw;
            }
        }

        private List<Question> ParseOpenQuestions(OleDbConnection connection)
        {
            var questions = new List<Question>();
            var correctAnswers = new Dictionary<int, HashSet<string>>();

            // Загружаем правильные ответы
            using (var command = new OleDbCommand("SELECT idTZ, answer FROM OpenOtvet", connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var questionId = reader.GetInt32(0);
                    var answer = reader.GetString(1);

                    if (!correctAnswers.ContainsKey(questionId))
                        correctAnswers[questionId] = new HashSet<string>();
                    correctAnswers[questionId].Add(answer);
                }
            }

            // Загружаем вопросы
            using (var command = new OleDbCommand("SELECT idTZ, text FROM OpenClose", connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var questionId = reader.GetInt32(0);
                    var text = reader.GetString(1);

                    var question = new Question
                    {
                        Id = questionId,
                        Text = text,
                        IsOpenQuestion = true,
                        CorrectAnswers = correctAnswers.ContainsKey(questionId)
                            ? correctAnswers[questionId].ToList()
                            : new List<string>(),
                        WrongAnswers = new List<string>()
                    };
                    questions.Add(question);
                }
            }

            return questions;
        }

        private List<Question> ParseClosedQuestions(OleDbConnection connection)
        {
            var questions = new Dictionary<int, Question>();

            using (var command = new OleDbCommand("SELECT idTZ, answer, [True] FROM CloseOtvet", connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    var questionId = reader.GetInt32(0);
                    var answer = reader.GetString(1);
                    var isCorrect = reader.GetBoolean(2);

                    if (!questions.ContainsKey(questionId))
                    {
                        questions[questionId] = new Question
                        {
                            Id = questionId,
                            IsOpenQuestion = false,
                            CorrectAnswers = new List<string>(),
                            WrongAnswers = new List<string>()
                        };
                    }

                    if (isCorrect)
                        questions[questionId].CorrectAnswers.Add(answer);
                    else
                        questions[questionId].WrongAnswers.Add(answer);
                }
            }

            // Загружаем тексты вопросов для закрытых вопросов
            foreach (var questionId in questions.Keys.ToList())
            {
                using (var command = new OleDbCommand("SELECT text FROM OpenClose WHERE idTZ = @id", connection))
                {
                    command.Parameters.AddWithValue("@id", questionId);
                    var text = command.ExecuteScalar() as string;
                    if (text != null)
                    {
                        questions[questionId].Text = text;
                    }
                }
            }

            return questions.Values.ToList();
        }

        public void PrintQuestionStats()
        {
            var questions = ParseQuestions();

            Console.WriteLine($"Всего вопросов: {questions.Count}");
            Console.WriteLine($"Открытых вопросов: {questions.Count(q => q.IsOpenQuestion)}");
            Console.WriteLine($"Закрытых вопросов: {questions.Count(q => !q.IsOpenQuestion)}");

            Console.WriteLine("\nСтатистика по правильным ответам:");
            var answerStats = questions
                .GroupBy(q => q.CorrectAnswers.Count)
                .OrderBy(g => g.Key);

            foreach (var group in answerStats)
            {
                Console.WriteLine($"Вопросов с {group.Key} правильными ответами: {group.Count()}");
            }

            Console.WriteLine("\nСтатистика по темам:");
            var topicStats = questions
                .GroupBy(q => _topics.ContainsKey(q.Id) ? _topics[q.Id] : "Без темы")
                .OrderBy(g => g.Key);

            foreach (var group in topicStats)
            {
                Console.WriteLine($"{group.Key}: {group.Count()} вопросов");
            }
        }

        public void ValidateAnswer(int questionId, string userAnswer)
        {
            try
            {
                EnsureConnectionOpen();

                // Проверяем открытые вопросы
                using (var command = new OleDbCommand(
                    "SELECT answer FROM OpenOtvet WHERE idTZ = @id", _connection))
                {
                    command.Parameters.AddWithValue("@id", questionId);
                    using (var reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var correctAnswer = reader.GetString(0);
                            if (string.Equals(userAnswer, correctAnswer, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine("Ответ верный!");
                                return;
                            }
                        }
                    }
                }

                // Проверяем закрытые вопросы
                using (var command = new OleDbCommand(
                    "SELECT [True] FROM CloseOtvet WHERE idTZ = @id AND answer = @answer", _connection))
                {
                    command.Parameters.AddWithValue("@id", questionId);
                    command.Parameters.AddWithValue("@answer", userAnswer);
                    var result = command.ExecuteScalar();

                    if (result != null && (bool)result)
                    {
                        Console.WriteLine("Ответ верный!");
                        return;
                    }
                }

                Console.WriteLine("Ответ неверный.");
            }
            catch (Exception ex) { Console.WriteLine("Поел говна в моменте: " + ex.Message); }
        }

        public List<Question> GetQuestionsByTopic(string topicName)
        {
            var questions = ParseQuestions();
            return questions.Where(q => _topics.ContainsKey(q.Id) && _topics[q.Id].Contains(topicName)).ToList();
        }
    }
}