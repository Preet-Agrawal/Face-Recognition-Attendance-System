Face Recognition Attendance System -

A real-time Face Recognition Attendance System built using Python, MTCNN, OpenCV, and face-recognition.
It detects faces through a webcam, recognizes registered users, and automatically marks attendance with secure SQL storage and Excel export.

use python app.py or python3 app.py to run in terminal
Features -

Real-time face detection using MTCNN
Face recognition & encoding with face-recognition
Automatic attendance marking with timestamps
SQL database integration for secure & persistent storage
One-click Excel export of attendance reports
Easy registration of new users with image capture
Simple and clean UI for smooth usage

Tech Stack -
Python
MTCNN
OpenCV
face-recognition (dlib)
SQL (SQLite/MySQL)
NumPy
Pandas / OpenPyXL

Project Structure
Face-Recognition-Attendance-System/
│── database/               # SQL database files
│── images/                 # Registered user face images
│── attendance/             # Exported Excel sheets
│── main.py                 # Main application script
│── detector.py             # MTCNN-based face detection
│── recognizer.py           # Face encoding & matching
│── utils.py                # Helper functions
│── requirements.txt
│── README.md

How to Run

Clone the repository:
git clone https://github.com/Preet-Agrawal/Face-Recognition-Attendance-System
Install dependencies:
pip install -r requirements.txt

Run the project:

python main.py
Attendance Export - 

Attendance logs can be exported to an Excel sheet directly from the SQL database with one click.

Contributions
Pull requests are welcome! Feel free to suggest improvements or report issues.



