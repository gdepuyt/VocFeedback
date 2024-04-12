using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace VocPoc
{

    public class TestData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }

        public int AnswerQ1 { get; set; }
        public string AnswerQ2 { get; set; }
        public int AnswerQ3 { get; set; }
        public string AnswerQ4 { get; set; }
        public int AnswerQ5 { get; set; }
        public int AnswerQ6 { get; set; }
        public string AnswerQ7 { get; set; }
        public int AnswerQ8 { get; set; }
        public string AnswerQ9 { get; set; }
        public int AnswerQ10 { get; set; }
    }

    internal class FakeClients
    {

        private static readonly string[] firstNames = { "John", "Paul", "Ringo", "George" };
        private static readonly string[] lastNames = { "Lennon", "McCartney", "Starr", "Harrison" };
        private static readonly string[] domains = { "gmail.com", "telenet.be", "live.com", "proximux.be", "voo.be" };
        private static readonly string[] responses = { "oui", "non", "peut être", "je ne sais pas", "à définir" };
        private static readonly string[] comments = { "Je suis tout à fait satisfaite.Je n ai rien eu comme difficultés, " +
                "tout s est déroulé dans la simplicité, sans heurt ni problème",
                "Hellemal good",
                "Vraiment bien",
                "Nul",
                "Quel dklsdqvlknqvlknqlk nsqdkln qkln lkn qkln mqlkn mknqm lkn mkln kqnd kln lmkq nmqkn mlkfq nmqkn lkqn" };

        public static List<TestData> GenerateClients(int numClients)
        {
           
            List<TestData> FakesurveysResults = new List<TestData>();

            for (int i = 0; i < numClients; i++)
            {
              //  var random = new Random();
                string firstName = firstNames[VocUtilsCommon.random.Next(0, firstNames.Length)];
                string lastName = lastNames[VocUtilsCommon.random.Next(0, lastNames.Length)];
                string domain = domains[VocUtilsCommon.random.Next(0, domains.Length)];
                string email = $"{firstName}.{lastName}@{domain}";

                FakesurveysResults.Add(new TestData
                {
                    FirstName = firstName,
                    LastName = lastName,
                    Email = email,
                    // You can add logic to generate random answers for each question here
                    AnswerQ1 = VocUtilsCommon.GetRandomNumber(0, 10),
                    AnswerQ2 = VocUtilsCommon.GetRandomItemFromArray(responses),
                    AnswerQ3 = VocUtilsCommon.GetRandomNumber(0, 10),
                    AnswerQ4 = VocUtilsCommon.GetRandomItemFromArray(comments),
                    AnswerQ5 = VocUtilsCommon.GetRandomNumber(0, 10),
                    AnswerQ6 = VocUtilsCommon.GetRandomNumber(0, 10),
                    AnswerQ7 = VocUtilsCommon.GetRandomItemFromArray(responses),
                    AnswerQ8 = VocUtilsCommon.GetRandomNumber(0, 10),
                    AnswerQ9 = VocUtilsCommon.GetRandomItemFromArray(responses),
                    AnswerQ10 = VocUtilsCommon.GetRandomNumber(0, 5)
                });
            }

            return FakesurveysResults;
        }
    }
}
