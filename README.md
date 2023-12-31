# DocGeneratorApp

## Overview
DocGeneratorApp is a user-friendly, PyQt5-based application designed for generating customized documents effortlessly. It leverages the `python-docx` library to manipulate Word documents, allowing users to create documents with various customizable inputs such as company details, meeting types, attendees, and more.

## Features
- **Customizable Document Generation**: Create documents based on user input, with the ability to insert logos, company names, meeting details, and more.
- **Dynamic Attendee List**: Add or modify the list of attendees for each document.
- **Template-based Design**: Utilizes a Word document template to ensure consistent formatting and structure.
- **User-Friendly Interface**: Easy-to-use graphical interface, making document creation accessible to all users.
- **Error Handling and Feedback**: Incorporates error handling and user feedback mechanisms for a smooth experience.

## Installation
To set up the DocGeneratorApp on your system, follow these steps:

1. Ensure that Python 3 and PyQt5 are installed on your computer. If not, you can download Python from https://www.python.org/downloads/ and install PyQt5 using pip:

   ```
   pip install pyqt5
   ```

2. Clone the repository or download the source code to your local machine.

3. Navigate to the application's directory in the terminal or command prompt.

4. Run the application using Python:

   ```
   python main.py
   ```

## Usage
Upon launching the DocGeneratorApp, you'll be greeted with a straightforward interface:

1. **Select or Upload Logo**: Choose a default logo or upload a custom one for your document.

2. **Company and Meeting Details**: Select a company from the predefined list or enter a custom name. Set the meeting type and fill in other relevant details.

3. **Manage Attendees**: The default attendee is set as the 'Chairman of the Board'. You can add more attendees as needed.

4. **Enter Meeting Information**: Input details like the meeting date, time, and addresses.

5. **Discussion and Resolutions**: Write down the key discussion points and resolutions in the respective fields.

6. **Generate Document**: Click on 'Generate Document' to create your Word document, which will be saved automatically with a unique filename.

## Contributing
We welcome contributions to improve DocGeneratorApp. If you have suggestions or enhancements, feel free to fork the repository and submit a pull request.

## Support
If you encounter any issues or have questions, please open an issue on the GitHub repository.

## License
This project is licensed under the MIT License - see the [LICENSE] file for details.
