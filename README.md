# HandWriter ✍🏽🐍

HandWriter is an advanced document processing application built with Python and PyQt5. It features a user-friendly graphical interface that allows users to convert `.docx` documents into handwritten-style PDFs with ease.

## Features 🔐🌱

### Core Functionalities:
- **Document Selection**: Easily select `.docx` files through a file dialog or drag-and-drop functionality.
- **Handwriting Conversion**: Convert `.docx` documents to handwritten PDFs with customizable settings.
- **Progress Indicators**: Real-time progress bar to track the conversion process.
- **Error Handling**: Pop-up notifications for missing files or unsupported characters.
- **Cross-Platform Support**: Works seamlessly on Windows, macOS, and Linux.

### User Interface Highlights:
- **Modern Design**: Minimalistic and intuitive interface styled with custom stylesheets.
- **Keyboard Shortcuts**: Quickly access core functionalities using shortcuts (e.g., `Ctrl+O` for opening files).
- **Animated Logo**: Dynamic branding using animated GIFs.
- **Responsive Buttons**: Button states visually adapt to user actions.

## How It Works

1. **Select a Document**:
   - Use the `Select Document` button or drag and drop a `.docx` file into the application.
2. **Review the Selection**:
   - The selected file name is displayed on the button.
3. **Start Conversion**:
   - Click the `Write` button to start converting the document into a handwritten-style PDF.
4. **View Output**:
   - Upon successful conversion, a message box displays the file's location with an option to open the containing folder.

## Technical Details

### Built With 🔗:
- **Python**
- **PyQt5**
- **docx Library**: For processing `.docx` documents.
- **joblib**: For managing font hashes.
- **fbs_runtime**: For packaging and resource management.

### Application Structure👨🏽‍💻:
- **UI_MainWindow**: Main application window and logic.
- **ParserThread**: Background thread for document processing.
- **MovieBox**: Custom utility for resizing GIFs used in the UI.

### File Structure:
```plaintext
.
├── main.py                 # Entry point of the application
├── document_parser.py      # Core document parsing logic
├── resources/              # Stylesheets and assets
│   ├── btn_select_unselected.qss
│   ├── btn_select_selected.qss
│   ├── btn_write_inactive.qss
│   ├── btn_write_active.qss
│   ├── progressbar_busy.qss
│   ├── progressbar_finished.qss
│   ├── DSC_logo_animated.gif
│   ├── handwriter_logo.png
│   ├── handwriter_logo_small.png
│   ├── hashes.pickle
```

## Installation

### Prerequisites:
- Python 3.8+
- Virtual Environment (recommended)

### Steps:
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/handwriter.git
   cd handwriter
   ```
2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate    # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the application:
   ```bash
   python main.py
   ```

## Screenshots

### Main Interface:
![Main Interface](https://github.com/imfallah/HandWritten-PyQt5/blob/main/hnd1.png)

### Conversion Progress:
![Conversion Progress](https://github.com/imfallah/HandWritten-PyQt5/blob/main/hnd2.gif)

## Contribution
Contributions are welcome! Please follow these steps:
1. Fork the repository.
2. Create a new branch (`feature/my-new-feature`).
3. Commit your changes (`git commit -am 'Add some feature'`).
4. Push to the branch (`git push origin feature/my-new-feature`).
5. Create a pull request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgements
- PyQt5 Documentation
- Python-docx Library


