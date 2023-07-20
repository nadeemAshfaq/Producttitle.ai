// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.

        const API_URL = 'https://api.openai.com/v1/chat/completions';
        const API_KEY = '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';

        function producttile(query1, query2, query3) {
            const data = {
                'model': 'gpt-3.5-turbo',
                'messages': [
                    {
                        'role': 'system',
                        'content': 'You are a helpful assistant.'
                    },
                    {
                        'role': 'user',
                        'content': "create short title based on product name: " + query1 + ' and description: ' + query2 + 'and Requsest' + query3
                    }
                ],
                'max_tokens': 50,
                'temperature': 0.5,
                n: 1
            };

            return new Promise((resolve, reject) => {
                $.ajax({
                    url: API_URL,
                    headers: {
                        'Authorization': `Bearer ${API_KEY}`,
                        'Content-Type': 'application/json'
                    },
                    method: 'POST',
                    dataType: 'json',
                    data: JSON.stringify(data),
                    success: function (response) {
                        /*console.log('ChatGPT response:', response);*/
                        const reply = response.choices[0].message.content;
                        console.log('Generated reply:', reply);
                        resolve(reply);
                    },
                    error: function (jqXHR, textStatus, errorThrown) {
                        console.error('AJAX request failed:', textStatus, errorThrown);
                        reject(new Error(errorThrown));
                    }
                });
            });
        }

        CustomFunctions.associate("PRODUCTTILE", producttile);
    }
})();


