# Bot-of-automation-using-selenium
WebDriver Initialization:

The script initializes a Selenium WebDriver (Chrome in headless mode) to load a chatbot page (https://accelirate.stagingbuilds.com/accelirate-assistant-chatbot/).

If the chatbot is embedded inside an iframe, it switches to that iframe to interact with it.

Question-Answer Interaction:

It fetches a list of questions from an Excel file (Website-Agent-Training-Questions.xlsx).

For each question, the script sends the question to the chatbot's input field, waits for the response, and extracts the answer.

If there are multiple answers, it selects the second one; otherwise, it takes the first answer.

The script includes retry logic in case the chatbot fails to respond.

Data Logging:

The script stores both the chatbot's answers and the timestamps when each answer was fetched in the Excel file.

It also calculates and logs the total time taken for processing all questions.

Page Reloading and Waiting:

After each question is processed, the page is reloaded, and the script waits for the chatbot to be ready again before continuing.

Final Output:

After processing all questions, the script saves the updated Excel file with columns for the chatbot’s answers, timestamps, and total time taken for the process.
