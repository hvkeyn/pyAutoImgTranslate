pyAutoImgTranslate is a utility that automatically translates text from screenshots through OCR and AI.

## Main functions

- Screenshot of the area: Activated by pressing Win+Shift+S to select the desired area of the screen.
- OCR with rectification: Text recognition with automatic tilt correction (deskew) using Tesseract and OpenCV.
- Translation: Sending the recognized text to the local LM Studio server for translation into Russian.
- User-friendly interface: Scalable result window with image and translation text, the ability to copy the selected fragment.
- Working in the background: The application runs in the system tray with a convenient exit.

## Technologies used

- Python  
- Tkinter – graphical interface  
- Tesseract OCR – text recognition  
- OpenCV & NumPy – image tilt correction  
- LM Studio API translation  
- Pystray – icon in the system tray  
- Keyboard & Mouse – global keyboard shortcuts

---

The project is designed for fast and effective translation of texts from the screen without unnecessary actions.