using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace TextForge
{
    internal class Ollama
    {
        private Uri _endpoint;
        private static readonly Uri _showApiRelativeUri = new Uri("api/show", UriKind.Relative);

        public Ollama(Uri endpoint)
        {
            _endpoint = endpoint;
        }

        public async Task<Dictionary<string, object>> Show(string modelName, bool verbose = false)
        {
            Uri fullUri = new Uri(_endpoint, _showApiRelativeUri);
                
            // Create the request payload
            var payload = new
            {
                name = modelName,
                verbose = verbose
            };

            // Serialize the payload to JSON using System.Text.Json
            var jsonPayload = JsonSerializer.Serialize(payload);

            // Create a StringContent object to send the payload
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            // Send the POST request
            var response = await CommonUtils.client.PostAsync(fullUri.OriginalString, content);
            response.EnsureSuccessStatusCode();

            // Read the response content as a string
            var responseString = await response.Content.ReadAsStringAsync();

            // Deserialize the JSON response into a Dictionary
            var responseData = JsonSerializer.Deserialize<Dictionary<string, object>>(responseString);

            return responseData;
        }
    }
}
