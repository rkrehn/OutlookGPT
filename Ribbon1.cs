using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using System.Net;
using EmailReplyParser;
using Microsoft.Office.Core;
using System.Configuration;
using System.IO;

namespace OutlookGPT
{
    public partial class Ribbon1
    {
        public string apikey = "";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // StreamWriter sw = new StreamWriter(Application.UserAppDataPath + "openaikey.dat"
            if (!File.Exists(System.Windows.Forms.Application.UserAppDataPath + "openaikey.dat"))
            {
                Form frm2 = new Form2();
                frm2.ShowDialog();
                return;
            }
            else
            {
                using (StreamReader sr = new StreamReader(System.Windows.Forms.Application.UserAppDataPath + "openaikey.dat"))
                {
                    while (!sr.EndOfStream)
                    {
                        apikey = sr.ReadLine();
                        break;
                    }
                }
            }
        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void dropDown1_ButtonClick(object sender, RibbonControlEventArgs e)
        {

        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // don't want them moving on without setting an API key
            if (apikey.Length < 5)
            {
                Form frm2 = new Form2();
                frm2.ShowDialog();
                return;
            }

            // find the compose window
            Outlook.Application application = Globals.ThisAddIn.Application;
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;

            // verify there is one
            if (mailItem == null)
            {
                MessageBox.Show("I couldn't find an email. Sorry.");
                return;
            }

            string prompt = "";
            string systemrole = "";
            // go through options
            switch (dropDown1.SelectedItem.Label)
            {
                case "Positive":
                    systemrole = "Please rephrase the following statement as a positive message: ";
                    break;
                case "Conscionable":
                    systemrole = "Please rephrase the following statement more conscionable: ";
                    break;
                case "Politically Correct":
                    systemrole = "Please rephrase the following statement as politically correct as possible: ";
                    break;
                case "Stern - Gently":
                    systemrole = "Please rephrase the following statement very stern, but also gently: ";
                    break;
                case "Stern - Direct":
                    systemrole = "Please rephrase the following statement very stern and direct: ";
                    break;
                case "Shorten":
                    systemrole = "Please rephrase the following statement so it's more succint and professional: ";
                    break;
            }

            // add stripped mail body
            string oldmail = mailItem.HTMLBody;
            var parsed_body = EmailReplyParser.EmailParser.Parse(mailItem.Body);
            prompt += parsed_body.Fragments[0].Content;

            //prepare all the parameters
            //string apiKey = "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";
            string model = "gpt-3.5-turbo";
            int maxTokens = 1024;
            float temperature = 0.7f;

            // Build the API request
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                string requestUrl = $"https://api.openai.com/v1/chat/completions";
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apikey}");

                var requestJson = new
                {
                    messages = new[]
                    {
                        new
                        {
                            role = "system",
                            content = systemrole
                        },
                        new
                        {
                            role = "user",
                            content = prompt
                        }
                    },
                    max_tokens = maxTokens,
                    temperature = temperature,
                    model = model
                };

                StringContent content = new StringContent(JsonConvert.SerializeObject(requestJson), Encoding.UTF8, "application/json");

                // Send the request and receive the response
                HttpResponseMessage response = client.PostAsync(requestUrl, content).Result;
                string responseJson = response.Content.ReadAsStringAsync().Result;

                // Extract the completed text from the response
                dynamic responseObject = JsonConvert.DeserializeObject(responseJson);
                string completedText = responseObject.choices[0].message.content;
                completedText = completedText.Replace("\n", "<br>");
                mailItem.HTMLBody = completedText + "<hr>" + oldmail;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.ToString() + Environment.NewLine + Environment.NewLine + "I am opening up the Open API form just in case.", "Outlook GPT Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Form frm = new Form2();
                frm.ShowDialog();
            }

        }

        private void btnKey_Click(object sender, RibbonControlEventArgs e)
        {
            Form frm2 = new Form2();
            frm2.ShowDialog();
        }
    }
}