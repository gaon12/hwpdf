import sys
import os
import re
import argparse
import pythoncom
import win32com.client

# ===== 공통 변환 로직 =====
def convert_hwp_to_pdf(folder_path=None, files=None, overwrite=False, log_func=print):
    """
    folder_path: 기본 폴더(없고 files가 절대경로면 필요 없음)
    files: 변환할 파일 경로 리스트(절대경로 가능). None이면 folder_path에서 스캔
    overwrite: 기존 PDF 덮어쓰기 여부
    log_func: 로그 출력용 콜백(print 같은 것)
    """
    pythoncom.CoInitialize()
    try:
        try:
            hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
        except Exception:
            hwp = win32com.client.Dispatch('HWPFrame.HwpObject')
        hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')

        if files is None:
            if not folder_path:
                log_func("폴더 또는 파일이 지정되지 않았습니다.")
                return
            files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                     if re.search(r"\.hwp(x?)$", f, re.IGNORECASE)]
        else:
            # 상대경로가 왔다면 folder_path 기준으로 절대경로화
            abs_files = []
            for f in files:
                abs_files.append(f if os.path.isabs(f) else os.path.join(folder_path or "", f))
            files = abs_files

        # 덮어쓰기 옵션 처리
        if not overwrite:
            files = [f for f in files
                     if not os.path.exists(os.path.splitext(f)[0] + '.pdf')]

        total = len(files)
        if total == 0:
            log_func("변환할 파일이 없습니다.")
            return

        for idx, input_path in enumerate(files, start=1):
            output_path = os.path.splitext(input_path)[0] + '.pdf'
            try:
                log_func(f"({idx}/{total}) 변환 시작: {os.path.basename(input_path)}")
                hwp.Open(input_path)
                hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
                hwp.HParameterSet.HFileOpenSave.filename = output_path
                hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
                log_func(f"({idx}/{total}) 변환 완료: {os.path.basename(input_path)}")
            except Exception as e:
                log_func(f"({idx}/{total}) 변환 실패: {os.path.basename(input_path)} -> {e}")

    finally:
        try:
            hwp.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


# ===== GUI 부분 =====
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QListWidget,
    QFileDialog,
    QMessageBox
)
from PyQt6.QtGui import QIcon

class ConversionThread(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self._interrupted = False

    def requestInterruption(self):
        self._interrupted = True
        super().requestInterruption()

    def run(self):
        pythoncom.CoInitialize()
        try:
            try:
                hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
            except Exception:
                hwp = win32com.client.Dispatch('HWPFrame.HwpObject')
            hwp.RegisterModule('FilePathCheckDLL', 'SecurityModule')

            total = len(self.file_paths)

            for idx, input_path in enumerate(self.file_paths, start=1):
                if self._interrupted or self.isInterruptionRequested():
                    self.log.emit("(중단) 변환 중단됨.")
                    break

                output_path = os.path.splitext(input_path)[0] + '.pdf'
                try:
                    self.log.emit(f"({idx}/{total}) 변환 시작: {os.path.basename(input_path)}")
                    hwp.Open(input_path)
                    hwp.HAction.GetDefault("FileSaveAs_S",
                                         hwp.HParameterSet.HFileOpenSave.HSet)
                    hwp.HParameterSet.HFileOpenSave.filename = output_path
                    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
                    hwp.HAction.Execute("FileSaveAs_S",
                                        hwp.HParameterSet.HFileOpenSave.HSet)
                    self.log.emit(f"({idx}/{total}) 변환 완료: {os.path.basename(input_path)}")
                except Exception as e:
                    self.log.emit(f"({idx}/{total}) 변환 실패: {os.path.basename(input_path)} -> {e}")

                percent = int(idx / total * 100)
                self.progress.emit(percent)
        finally:
            try:
                hwp.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()
            self.finished.emit()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWP/HWPX → PDF 변환기")

        # 아이콘 설정
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setAcceptDrops(True)
        self.setMinimumSize(540, 500)
        self.interrupted = False

        self.setStyleSheet("""
            QWidget { background-color: #f5f5f5; }
            QPushButton { background-color: #2196F3; color: white; border: none; padding: 10px 20px; border-radius: 12px; font-size: 14px; }
            QPushButton#start { background-color: #4CAF50; }
            QPushButton#stop { background-color: #f44336; }
            QPushButton:disabled { background-color: #BDBDBD; color: #757575; }
            QProgressBar { background-color: #E0E0E0; border-radius: 8px; text-align: center; height: 20px; }
            QProgressBar::chunk { border-radius: 8px; background-color: #4CAF50; }
            QListWidget { background-color: white; border: 1px solid #CCC; border-radius: 8px; padding: 5px; }
            QScrollBar:vertical { background: #f5f5f5; width: 12px; margin: 0px; border-radius: 6px; }
            QScrollBar::handle:vertical { background: #CCCCCC; min-height: 20px; border-radius: 6px; }
            QScrollBar::handle:vertical:hover { background: #BBBBBB; }
            QScrollBar::sub-line:vertical, QScrollBar::add-line:vertical { height: 0px; }
            QLabel { font-size: 14px; }
        """)

        self.info_label = QLabel("파일을 드래그하거나 '열기' 버튼으로 선택하세요.")
        self.open_btn = QPushButton("열기(파일 선택)")
        self.open_btn.clicked.connect(self.open_files)

        self.start_btn = QPushButton("변환 시작")
        self.start_btn.setObjectName("start")
        self.start_btn.clicked.connect(self.start_conversion)
        self.start_btn.setEnabled(False)

        self.stop_btn = QPushButton("변환 중단")
        self.stop_btn.setObjectName("stop")
        self.stop_btn.clicked.connect(self.stop_conversion)
        self.stop_btn.setEnabled(False)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        self.log_list = QListWidget()

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.open_btn)
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.setSpacing(10)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.info_label)
        main_layout.addLayout(btn_layout)
        main_layout.addWidget(self.progress_bar)
        main_layout.addWidget(self.log_list)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(12)

        self.selected_files = []  # 절대경로 리스트
        self.thread = None

        # 로그 자동 스크롤
        self.log_list.model().rowsInserted.connect(lambda: self.log_list.scrollToBottom())

    # ----- Drag & Drop -----
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        if self.thread_running():
            return
        urls = event.mimeData().urls()
        if not urls:
            return
        paths = [u.toLocalFile() for u in urls]
        files = []
        for p in paths:
            if os.path.isdir(p):
                files.extend([os.path.join(p, f) for f in os.listdir(p)
                              if re.search(r"\.hwp(x?)$", f, re.IGNORECASE)])
            elif os.path.isfile(p) and re.search(r"\.hwp(x?)$", p, re.IGNORECASE):
                files.append(p)
        if files:
            self.set_files(files)

    # ----- File Dialog -----
    def open_files(self):
        if self.thread_running():
            return
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "HWP/HWPX 파일 선택",
            "",
            "HWP Files (*.hwp *.hwpx);;All Files (*)"
        )
        if not files:
            return

        # 서로 다른 폴더 섞임 방지(원하면 제거 가능)
        base_dir = os.path.dirname(files[0])
        if any(os.path.dirname(f) != base_dir for f in files):
            QMessageBox.warning(self, "경고", "같은 폴더의 파일만 선택하세요.")
            return

        self.set_files(files)

    def set_files(self, files):
        # 파일 리스트 설정
        self.selected_files = files
        self.info_label.setText(f"선택된 파일 {len(files)}개")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_list.clear()

    # ----- Conversion -----
    def start_conversion(self):
        if not self.selected_files:
            return

        # 중복 PDF 체크
        duplicates = [f for f in self.selected_files
                      if os.path.exists(os.path.splitext(f)[0] + '.pdf')]
        file_list_names = [os.path.basename(f) for f in duplicates]

        files_to_convert = self.selected_files[:]
        if duplicates:
            msg = "다음 파일의 PDF가 이미 존재합니다:\n" + "\n".join(file_list_names)
            dlg = QMessageBox(self)
            dlg.setWindowTitle("덮어쓰기 확인")
            dlg.setText(msg)
            overwriteAllBtn = dlg.addButton("모두 덮어쓰기", QMessageBox.ButtonRole.AcceptRole)
            skipBtn = dlg.addButton("건너뛰기", QMessageBox.ButtonRole.RejectRole)
            cancelBtn = dlg.addButton("취소", QMessageBox.ButtonRole.DestructiveRole)
            dlg.exec()
            clicked = dlg.clickedButton()
            if clicked == cancelBtn:
                return
            elif clicked == skipBtn:
                files_to_convert = [f for f in files_to_convert if f not in duplicates]
            # 모두 덮어쓰기면 그대로

        self.interrupted = False
        self.open_btn.setEnabled(False)
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        self.thread = ConversionThread(files_to_convert)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.log.connect(self.log_list.addItem)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def stop_conversion(self):
        if self.thread and self.thread.isRunning():
            self.interrupted = True
            self.thread.requestInterruption()
            self.stop_btn.setEnabled(False)

    def on_finished(self):
        if self.interrupted:
            self.info_label.setText("변환이 중단되었습니다.")
        else:
            self.info_label.setText("모든 파일이 성공적으로 변환되었습니다.")
        self.open_btn.setEnabled(True)
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    def thread_running(self):
        return self.thread is not None and self.thread.isRunning()


# ===== 엔트리 포인트 =====
def main():
    parser = argparse.ArgumentParser(description="HWP/HWPX → PDF 변환기 (GUI/CLI 겸용)")
    parser.add_argument("-d", "--folder", help="변환할 폴더 경로")
    parser.add_argument("-f", "--files", nargs="*", help="변환할 파일 경로(여러 개 가능, 절대/상대 혼용 가능)")
    parser.add_argument("--overwrite", action="store_true", help="기존 PDF가 있어도 덮어쓰기")
    parser.add_argument("--gui", action="store_true", help="강제로 GUI 실행")
    args = parser.parse_args()

    # CLI 모드: 폴더나 파일이 주어졌고 --gui가 아니면
    if (args.folder or args.files) and not args.gui:
        # 절대경로 정리
        files = None
        if args.files:
            files = [os.path.abspath(p) for p in args.files]
        folder = os.path.abspath(args.folder) if args.folder else None

        # 유효성 체크
        if folder and not os.path.isdir(folder):
            print("폴더 경로가 올바르지 않습니다.", file=sys.stderr)
            sys.exit(1)
        convert_hwp_to_pdf(folder_path=folder, files=files, overwrite=args.overwrite, log_func=print)
        sys.exit(0)

    # GUI 실행
    app = QApplication(sys.argv)
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
