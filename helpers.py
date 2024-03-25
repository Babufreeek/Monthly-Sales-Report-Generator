from PyQt5.QtWidgets import QDialog, QVBoxLayout, QProgressBar
from PyQt5.QtCore import QThread, pyqtSignal, QTimer

class LoadingScreen(QDialog):
    """
    Loading screen
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Loading...")
        self.setFixedSize(300, 100)
        
        layout = QVBoxLayout(self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setStyleSheet("QProgressBar {"
                                         "border: 2px solid grey;"
                                         "border-radius: 5px;"
                                         "text-align: center;"
                                         "color: white;"
                                         "padding: 1px;"
                                         "background-color: #2196F3;"
                                         "}")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setRange(0, 0)  # Set to infinite progress
        layout.addWidget(self.progress_bar)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress)
        self.timer.start(50)  # Update progress every 50 milliseconds
        self.direction = 1  # 1 for forward, -1 for backward
        self.position = 0
     
    def update_progress(self):
        if self.direction == 1:
            self.position += 1
        else:
            self.position -= 1
        
        if self.position >= 100:
            self.direction = -1
        elif self.position <= 0:
            self.direction = 1
        
        self.progress_bar.setValue(self.position)

class FileProcessor(QThread):
    """
    Processor which performs some task and sends a signal when complete
    """
    finished = pyqtSignal()

    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        # Run function
        self.function(*self.args, **self.kwargs)
        
        self.finished.emit()
