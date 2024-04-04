# Multiple-Pdf-Excel-chat-application

## Introduction
Demonstrate the capabilities of OpenAI's chatbot in handling multiple Excel and PDF files for information extraction and manipulation.

## Setup Instructions

### Step 1: Clone the Repository
Clone the repository to your local machine:

```bash
git clone https://github.com/Sanjay3739/Multiple-Pdf-excel-chat-application.git	
```

### Step 2: Create a Virtual Environment
Navigate to the project directory and create a virtual environment using Python's venv module:

```bash
python -m venv venv
```

### Step 3: Activate the Virtual Environment
Activate the virtual environment:

```bash
venv\Scripts\activate  # For Windows
source venv/bin/activate  # For Unix/Mac
```

### Step 4: Install Dependencies
Install the required dependencies listed in `requirements.txt`:

```bash
pip install -r requirements.txt
```

### Step 5: Obtain an API Key from OpenAI and Add it to the .env File
Obtain an API key from OpenAI and add it to the `.env` file in the project directory:

```commandline
OPENAI_API_KEY=your_secret_api_key
```

### Step 6: Run the Application
Run the Streamlit application:

```bash
streamlit run app.py
```

## Usage
Once the application is running, access it through your web browser. Follow the on-screen instructions to use the chat interface for generating PDF documents.

## Additional Notes
- Make sure to provide appropriate inputs and follow the prompts within the Streamlit interface.
- Customize the functionality as needed based on your project requirements.
