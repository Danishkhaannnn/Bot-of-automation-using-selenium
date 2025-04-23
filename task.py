import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from datetime import datetime  # Importing datetime to get the current time

# Function to initialize the WebDriver and load the page
def initialize_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (no UI)
    chrome_options.add_argument("--no-sandbox")  # Disable sandboxing for Docker environments
    chrome_options.add_argument("--disable-dev-shm-usage")  # For environments with limited resources
    
    service = Service(ChromeDriverManager().install())  # Automatically get the correct chromedriver version
    driver = webdriver.Chrome(service=service)

    driver.get('https://accelirate.stagingbuilds.com/accelirate-assistant-chatbot/')
    
    # Check if the chatbot is inside an iframe and switch to it if necessary
    try:
        iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)
        print("Switched to iframe.")
    except Exception as e:
        print(f"No iframe found, continuing without iframe: {e}")

    # Wait for the input field to load (prompt)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, '//input[@class="webchat__send-box-text-box__input" and @placeholder="Type your message"]')))
    return driver

# Function to get the answer for a question with retry logic
def get_answer(driver, question, retries=3):
    while retries > 0:
        try:
            # Wait for the input field to be clickable
            input_field = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//input[@class="webchat__send-box-text-box__input" and @placeholder="Type your message"]')))
            input_field.clear()
            input_field.send_keys(question)  # Type the question
            input_field.send_keys(Keys.RETURN)  # Press 'Enter' to submit the question

            # Sleep to ensure the answer has time to load
            time.sleep(30)  # Adjusted sleep time

            # Fetch all the possible answers
            answer_containers = WebDriverWait(driver, 90).until(EC.presence_of_all_elements_located((By.XPATH, '//div[contains(@class, "webchat__text-content--is-markdown")]')))

            # Fetch the second answer if multiple answers exist, otherwise the first one
            if len(answer_containers) > 1:
                answer = answer_containers[1].text  # Take the second one
            else:
                answer = answer_containers[0].text  # If only one answer exists

            # Debug: Print the question and answer
            print(f"Question: {question}")
            print(f"Answer: {answer}")

            return answer
        except Exception as e:
            print(f"Error fetching answer for '{question}': {e}")
            retries -= 1
            if retries == 0:
                return 'No answer found'
            time.sleep(5)  # Retry after a brief pause

# Load the new Excel file with questions, answers, and chatbot responses
df = pd.read_excel('Website-Agent-Training-Questions.xlsx')  # Adjust to the uploaded file

# Check if the 'chatbot' and 'timestamp' columns exist, if not, create them
if 'chatbot' not in df.columns:
    df['chatbot'] = ''  # Add an empty 'chatbot' column if it doesn't exist
if 'timestamp' not in df.columns:
    df['timestamp'] = ''  # Add an empty 'timestamp' column if it doesn't exist

# Initialize the driver once and keep it open
driver = initialize_driver()

# Record the start time
start_time = datetime.now()

# Loop through each question in the CSV file and get the answer
for index, row in df.iterrows():
    question = row['Question']  # 'Question' is the column containing the questions
    
    # Check if the question is valid (not NaN or empty)
    if pd.isna(question) or question.strip() == "":
        print(f"Skipping empty or invalid question at index {index}.")
        continue  # Skip empty or invalid questions
    
    print(f"Processing Question: {question}")
    
    # Get the answer for the question
    answer = get_answer(driver, question)
    
    # Get the current time and format it
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Time format: YYYY-MM-DD HH:MM:SS

    # Store the answer and timestamp in the dataframe
    df.at[index, 'chatbot'] = answer  # Add chatbot's answer
    df.at[index, 'timestamp'] = current_time  # Add timestamp for when the answer is saved

    print(f"Stored Answer: {answer} at {current_time}")

    # After every question, reload the page for the next question
    driver.get('https://accelirate.stagingbuilds.com/accelirate-assistant-chatbot/')
    try:
        iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)
        print("Switched to iframe after page reload.")
    except Exception as e:
        print(f"No iframe found after reload: {e}")
    
    # Wait a moment before continuing with the next question
    time.sleep(5)

# Record the end time
end_time = datetime.now()

# Calculate the total duration
total_duration = end_time - start_time
print(f"Total time taken to answer all questions: {total_duration}")

# Add the total time to the new 'total_time' column after the timestamp column
df['total_time'] = ''
df.at[len(df)-1, 'total_time'] = str(total_duration)  # Save the total duration in the total_time column

# Save the updated dataframe with answers, timestamps, and total time back to the same Excel file
df.to_excel('Website-Agent-Training-Questions.xlsx', index=False)

# Close the browser
driver.quit()
