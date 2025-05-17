#!/usr/bin/env python3
import sys
import os
import json
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTextEdit, QToolBar, QAction, QFileDialog, 
                            QMessageBox, QListWidget, QListWidgetItem, QComboBox,
                            QFontComboBox, QColorDialog, QDialog, QGridLayout,
                            QFrame, QSplitter, QProgressBar, QScrollArea,
                            QSizePolicy, QSpacerItem, QStyle, QStyleFactory,
                            QGroupBox)
from PyQt5.QtGui import (QIcon, QFont, QColor, QTextCharFormat, QTextCursor, 
                         QPalette, QPixmap, QTextListFormat, QTextFormat)
from PyQt5.QtCore import Qt, QSize, QPropertyAnimation, QEasingCurve, QRect, QTimer

class ModernButton(QPushButton):
    """Custom button with modern styling"""
    def __init__(self, text, parent=None, primary=False):
        super().__init__(text, parent)
        self.primary = primary
        self.setMinimumHeight(36)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(self._get_style())
        
    def _get_style(self):
        if self.primary:
            return """
                QPushButton {
                    background-color: #4a86e8;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    padding: 8px 16px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #3a76d8;
                }
                QPushButton:pressed {
                    background-color: #2a66c8;
                }
                QPushButton:disabled {
                    background-color: #cccccc;
                    color: #888888;
                }
            """
        else:
            return """
                QPushButton {
                    background-color: #f0f0f0;
                    color: #333333;
                    border: 1px solid #cccccc;
                    border-radius: 4px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                    border: 1px solid #bbbbbb;
                }
                QPushButton:pressed {
                    background-color: #d0d0d0;
                }
                QPushButton:disabled {
                    background-color: #f8f8f8;
                    color: #bbbbbb;
                    border: 1px solid #dddddd;
                }
            """

class TextFormatButton(QPushButton):
    """Custom button for text formatting with modern styling"""
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(28)
        self.setMinimumWidth(80)
        self.setCursor(Qt.PointingHandCursor)
        self.setCheckable(True)
        self.setStyleSheet("""
            QPushButton {
                background-color: #f8f8f8;
                color: #333333;
                border: 1px solid #dddddd;
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
                border: 1px solid #cccccc;
            }
            QPushButton:pressed, QPushButton:checked {
                background-color: #4a86e8;
                color: white;
                border: 1px solid #3a76d8;
            }
            QPushButton:disabled {
                background-color: #f8f8f8;
                color: #bbbbbb;
                border: 1px solid #dddddd;
            }
        """)

class ModernLineEdit(QLineEdit):
    """Custom line edit with modern styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(36)
        self.setStyleSheet("""
            QLineEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #4a86e8;
            }
            QLineEdit:disabled {
                background-color: #f0f0f0;
                color: #888888;
            }
        """)

class ValidationPanel(QWidget):
    """Panel to display validation results with integrated action buttons"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        
        # Main layout
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(12, 12, 12, 12)
        self.layout.setSpacing(10)
        
        # Title
        title_label = QLabel("Validation Results")
        title_label.setStyleSheet("""
            font-size: 16px;
            font-weight: bold;
            color: #4a86e8;
            padding-bottom: 8px;
        """)
        self.layout.addWidget(title_label)
        
        # Result area
        self.result_area = QTextEdit()
        self.result_area.setReadOnly(True)
        self.result_area.setMinimumHeight(150)
        self.result_area.setStyleSheet("""
            QTextEdit {
                border: 1px solid #dddddd;
                border-radius: 4px;
                background-color: white;
                padding: 8px;
            }
        """)
        self.layout.addWidget(self.result_area)
        
        # Action buttons (initially hidden)
        self.action_frame = QFrame()
        self.action_frame.setVisible(False)
        self.action_layout = QHBoxLayout(self.action_frame)
        self.action_layout.setContentsMargins(0, 0, 0, 0)
        
        self.edit_button = ModernButton("Edit Email", primary=False)
        self.edit_button.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogResetButton))
        self.edit_button.clicked.connect(self.on_edit)
        
        self.send_button = ModernButton("Send Anyway", primary=True)
        self.send_button.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogApplyButton))
        self.send_button.clicked.connect(self.on_send)
        
        self.abort_button = ModernButton("Abort", primary=False)
        self.abort_button.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogCancelButton))
        self.abort_button.clicked.connect(self.on_abort)
        
        self.action_layout.addWidget(self.edit_button)
        self.action_layout.addWidget(self.send_button)
        self.action_layout.addWidget(self.abort_button)
        
        self.layout.addWidget(self.action_frame)
        
        # Refine button (initially hidden)
        self.refine_button = ModernButton("Refine Email", primary=True)
        self.refine_button.clicked.connect(self.on_refine)
        self.refine_button.setVisible(False)
        self.layout.addWidget(self.refine_button)
        
    def set_result(self, result):
        self.result_area.setHtml(result)
        
    def show_actions(self, show=True):
        self.action_frame.setVisible(show)
        
    def show_refine_button(self, show=True):
        self.refine_button.setVisible(show)
        
    def on_edit(self):
        self.show_actions(False)
        # Just continue editing
        
    def on_send(self):
        self.show_actions(False)
        self.parent.send_email()
        
    def on_abort(self):
        self.show_actions(False)
        self.parent.clear_form()
        
    def on_refine(self):
        self.parent.refine_email()

class RefinedContentPanel(QWidget):
    """Panel to display refined email content with insert button"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        
        # Main layout
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(12, 12, 12, 12)
        self.layout.setSpacing(10)
        
        # Title
        title_label = QLabel("Refined Email Content")
        title_label.setStyleSheet("""
            font-size: 16px;
            font-weight: bold;
            color: #4a86e8;
            padding-bottom: 8px;
        """)
        self.layout.addWidget(title_label)
        
        # Content area
        self.content_area = QTextEdit()
        self.content_area.setReadOnly(True)
        self.content_area.setMinimumHeight(150)
        self.content_area.setStyleSheet("""
            QTextEdit {
                border: 1px solid #b0d0ff;
                border-radius: 4px;
                background-color: white;
                padding: 8px;
            }
        """)
        self.layout.addWidget(self.content_area)
        
        # Insert button
        self.insert_button = ModernButton("Insert Refined Content", primary=True)
        self.insert_button.clicked.connect(self.on_insert)
        self.layout.addWidget(self.insert_button)
        
        # Initially hide this panel
        self.setVisible(False)
        
    def set_content(self, content):
        self.content_area.setHtml(content)
        
    def on_insert(self):
        self.parent.insert_refined_content()

class AttachmentPanel(QFrame):
    """Compact attachment panel with modern styling"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setFrameShape(QFrame.StyledPanel)
        self.setFrameShadow(QFrame.Raised)
        self.setMaximumHeight(120)
        self.setStyleSheet("""
            AttachmentPanel {
                background-color: #f8f8f8;
                border: 1px solid #dddddd;
                border-radius: 4px;
            }
        """)
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(8, 8, 8, 8)
        
        # Header with label and buttons
        header_layout = QHBoxLayout()
        
        attachment_label = QLabel("Attachments:")
        attachment_label.setStyleSheet("font-weight: bold;")
        header_layout.addWidget(attachment_label)
        
        header_layout.addStretch()
        
        add_btn = ModernButton("Add", primary=False)
        add_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        add_btn.setToolTip("Add attachment")
        add_btn.clicked.connect(self.parent.add_attachment)
        add_btn.setMaximumWidth(80)
        
        remove_btn = ModernButton("Remove", primary=False)
        remove_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_TrashIcon))
        remove_btn.setToolTip("Remove selected attachment")
        remove_btn.clicked.connect(self.parent.remove_attachment)
        remove_btn.setMaximumWidth(80)
        
        header_layout.addWidget(add_btn)
        header_layout.addWidget(remove_btn)
        
        self.layout.addLayout(header_layout)
        
        # Attachment list
        self.attachment_list = QListWidget()
        self.attachment_list.setMaximumHeight(60)
        self.attachment_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #dddddd;
                border-radius: 4px;
                background-color: white;
            }
            QListWidget::item {
                padding: 4px;
            }
            QListWidget::item:selected {
                background-color: #e0e0e0;
                color: #333333;
            }
        """)
        self.layout.addWidget(self.attachment_list)

class CompositionPanel(QWidget):
    """Panel for email composition"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        
        # Main layout
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(16, 16, 16, 16)
        self.layout.setSpacing(12)
        
        # Header with app title
        header_label = QLabel("Email Composer")
        header_label.setStyleSheet("""
            font-size: 18px;
            font-weight: bold;
            color: #4a86e8;
            padding-bottom: 8px;
            border-bottom: 1px solid #dddddd;
        """)
        self.layout.addWidget(header_label)
        
        # Recipient section
        recipient_layout = QHBoxLayout()
        recipient_label = QLabel("To:")
        recipient_label.setMinimumWidth(60)
        self.recipient_input = ModernLineEdit()
        recipient_layout.addWidget(recipient_label)
        recipient_layout.addWidget(self.recipient_input)
        self.layout.addLayout(recipient_layout)
        
        # Subject section
        subject_layout = QHBoxLayout()
        subject_label = QLabel("Subject:")
        subject_label.setMinimumWidth(60)
        self.subject_input = ModernLineEdit()
        subject_layout.addWidget(subject_label)
        subject_layout.addWidget(self.subject_input)
        self.layout.addLayout(subject_layout)
        
        # Create formatting toolbar
        self.create_formatting_toolbar()
        
        # Text editor
        self.text_editor = QTextEdit()
        self.text_editor.setMinimumHeight(250)
        self.text_editor.setStyleSheet("""
            QTextEdit {
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
            }
            QTextEdit:focus {
                border: 1px solid #4a86e8;
            }
        """)
        self.layout.addWidget(self.text_editor)
        
        # Attachments panel (compact)
        self.attachment_panel = AttachmentPanel(self.parent)
        self.layout.addWidget(self.attachment_panel)
        
        # Action buttons
        action_layout = QHBoxLayout()
        
        # Add spacer to push buttons to the right
        action_layout.addStretch()
        
        validate_btn = ModernButton("Validate with Gemini", primary=True)
        validate_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_DialogApplyButton))
        validate_btn.clicked.connect(self.parent.validate_email)
        validate_btn.setMinimumWidth(180)
        action_layout.addWidget(validate_btn)
        
        send_btn = ModernButton("Send Email", primary=True)
        send_btn.setIcon(QApplication.style().standardIcon(QStyle.SP_CommandLink))
        send_btn.clicked.connect(self.parent.confirm_send)
        send_btn.setMinimumWidth(150)
        action_layout.addWidget(send_btn)
        
        self.layout.addLayout(action_layout)
        
    def create_formatting_toolbar(self):
        # Create toolbar with modern styling
        self.formatting_toolbar = QToolBar("Formatting")
        self.formatting_toolbar.setMovable(False)
        self.formatting_toolbar.setStyleSheet("""
            QToolBar {
                background-color: #f8f8f8;
                border: 1px solid #dddddd;
                border-radius: 4px;
                spacing: 4px;
                padding: 4px;
            }
        """)
        self.parent.addToolBar(self.formatting_toolbar)
        
        # Font family
        font_family_label = QLabel("Font:")
        font_family_label.setStyleSheet("padding-left: 4px; padding-right: 4px;")
        self.formatting_toolbar.addWidget(font_family_label)
        
        self.font_family = QFontComboBox()
        self.font_family.setMaximumWidth(150)
        self.font_family.currentFontChanged.connect(self.parent.change_font_family)
        self.formatting_toolbar.addWidget(self.font_family)
        
        # Font size
        font_size_label = QLabel("Size:")
        font_size_label.setStyleSheet("padding-left: 8px; padding-right: 4px;")
        self.formatting_toolbar.addWidget(font_size_label)
        
        self.font_size = QComboBox()
        self.font_size.setMaximumWidth(60)
        font_sizes = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '22', '24', '26', '28', '36', '48', '72']
        self.font_size.addItems(font_sizes)
        self.font_size.setCurrentText('12')
        self.font_size.currentTextChanged.connect(self.parent.change_font_size)
        self.formatting_toolbar.addWidget(self.font_size)
        
        self.formatting_toolbar.addSeparator()
        
        # Text style buttons
        self.bold_button = TextFormatButton("Bold")
        self.bold_button.setToolTip("Bold (Ctrl+B)")
        self.bold_button.clicked.connect(self.parent.toggle_bold)
        self.formatting_toolbar.addWidget(self.bold_button)
        
        self.italic_button = TextFormatButton("Italic")
        self.italic_button.setToolTip("Italic (Ctrl+I)")
        self.italic_button.clicked.connect(self.parent.toggle_italic)
        self.formatting_toolbar.addWidget(self.italic_button)
        
        self.underline_button = TextFormatButton("Underline")
        self.underline_button.setToolTip("Underline (Ctrl+U)")
        self.underline_button.clicked.connect(self.parent.toggle_underline)
        self.formatting_toolbar.addWidget(self.underline_button)
        
        self.formatting_toolbar.addSeparator()
        
        # Color buttons
        self.text_color_button = TextFormatButton("Text Color")
        self.text_color_button.setToolTip("Change text color")
        self.text_color_button.clicked.connect(self.parent.change_text_color)
        self.formatting_toolbar.addWidget(self.text_color_button)
        
        self.bg_color_button = TextFormatButton("Highlight")
        self.bg_color_button.setToolTip("Change background color")
        self.bg_color_button.clicked.connect(self.parent.change_background_color)
        self.formatting_toolbar.addWidget(self.bg_color_button)
        
        self.formatting_toolbar.addSeparator()
        
        # Alignment buttons
        self.align_left_button = TextFormatButton("Left")
        self.align_left_button.setToolTip("Align text left")
        self.align_left_button.clicked.connect(lambda: self.parent.text_editor.setAlignment(Qt.AlignLeft))
        self.formatting_toolbar.addWidget(self.align_left_button)
        
        self.align_center_button = TextFormatButton("Center")
        self.align_center_button.setToolTip("Align text center")
        self.align_center_button.clicked.connect(lambda: self.parent.text_editor.setAlignment(Qt.AlignCenter))
        self.formatting_toolbar.addWidget(self.align_center_button)
        
        self.align_right_button = TextFormatButton("Right")
        self.align_right_button.setToolTip("Align text right")
        self.align_right_button.clicked.connect(lambda: self.parent.text_editor.setAlignment(Qt.AlignRight))
        self.formatting_toolbar.addWidget(self.align_right_button)
        
        self.align_justify_button = TextFormatButton("Justify")
        self.align_justify_button.setToolTip("Justify text")
        self.align_justify_button.clicked.connect(lambda: self.parent.text_editor.setAlignment(Qt.AlignJustify))
        self.formatting_toolbar.addWidget(self.align_justify_button)
        
        self.formatting_toolbar.addSeparator()
        
        # List buttons
        self.bullet_list_button = TextFormatButton("Bullet List")
        self.bullet_list_button.setToolTip("Insert bullet list")
        self.bullet_list_button.clicked.connect(self.parent.insert_bullet_list)
        self.formatting_toolbar.addWidget(self.bullet_list_button)
        
        self.numbered_list_button = TextFormatButton("Numbered")
        self.numbered_list_button.setToolTip("Insert numbered list")
        self.numbered_list_button.clicked.connect(self.parent.insert_numbered_list)
        self.formatting_toolbar.addWidget(self.numbered_list_button)

class EmailComposer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Composer with Gemini Validation")
        self.setGeometry(100, 100, 1200, 900)  # Increased window size
        
        # Set application style
        QApplication.setStyle(QStyleFactory.create("Fusion"))
        
        # Email credentials
        self.email = "Your email here"
        self.password = "App Password Here"
        
        # Gemini API key
        self.api_key = "API KEY HERE"
        
        # Initialize UI
        self.init_ui()
        
        # List to store attachments
        self.attachments = []
        
        # Store refined content
        self.refined_subject = ""
        self.refined_body_html = ""
        
        # Set window icon
        self.setWindowIcon(QApplication.style().standardIcon(QStyle.SP_MessageBoxInformation))
        
    def init_ui(self):
        # Set application palette for consistent colors
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(248, 248, 248))
        palette.setColor(QPalette.WindowText, QColor(51, 51, 51))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        palette.setColor(QPalette.ToolTipText, QColor(51, 51, 51))
        palette.setColor(QPalette.Text, QColor(51, 51, 51))
        palette.setColor(QPalette.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ButtonText, QColor(51, 51, 51))
        palette.setColor(QPalette.Highlight, QColor(74, 134, 232))
        palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
        QApplication.setPalette(palette)
        
        # Central widget with margin
        central_widget = QWidget()
        central_widget.setContentsMargins(16, 16, 16, 16)
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(16)
        
        # Create main splitter
        self.main_splitter = QSplitter(Qt.Vertical)
        self.main_splitter.setChildrenCollapsible(False)
        main_layout.addWidget(self.main_splitter)
        
        # Top panel - Composition
        self.composition_panel = CompositionPanel(self)
        self.main_splitter.addWidget(self.composition_panel)
        
        # Bottom panel - Results (with horizontal split)
        self.results_widget = QWidget()
        self.results_layout = QHBoxLayout(self.results_widget)
        self.results_layout.setContentsMargins(0, 0, 0, 0)
        self.results_layout.setSpacing(16)
        
        # Create horizontal splitter for results
        self.results_splitter = QSplitter(Qt.Horizontal)
        self.results_splitter.setChildrenCollapsible(False)
        self.results_layout.addWidget(self.results_splitter)
        
        # Left side - Validation results
        self.validation_panel = ValidationPanel(self)
        self.results_splitter.addWidget(self.validation_panel)
        
        # Right side - Refined content
        self.refined_panel = RefinedContentPanel(self)
        self.results_splitter.addWidget(self.refined_panel)
        
        # Add results widget to main splitter
        self.main_splitter.addWidget(self.results_widget)
        
        # Set initial splitter sizes (60% top, 40% bottom)
        self.main_splitter.setSizes([600, 400])
        self.results_splitter.setSizes([500, 500])
        
        # Status bar for notifications
        self.statusBar().showMessage("Ready")
        
        # Reference to text editor for convenience
        self.text_editor = self.composition_panel.text_editor
        
        # Set tab order for navigation
        self.setTabOrder(self.composition_panel.recipient_input, self.composition_panel.subject_input)
        self.setTabOrder(self.composition_panel.subject_input, self.text_editor)
        
    def change_font_family(self, font):
        self.text_editor.setCurrentFont(font)
        
    def change_font_size(self, size):
        self.text_editor.setFontPointSize(float(size))
        
    def toggle_bold(self, checked):
        if checked:
            self.text_editor.setFontWeight(QFont.Bold)
        else:
            self.text_editor.setFontWeight(QFont.Normal)
            
    def toggle_italic(self, checked):
        self.text_editor.setFontItalic(checked)
        
    def toggle_underline(self, checked):
        self.text_editor.setFontUnderline(checked)
        
    def change_text_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.text_editor.setTextColor(color)
            
    def change_background_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            cursor = self.text_editor.textCursor()
            format = QTextCharFormat()
            format.setBackground(color)
            cursor.mergeCharFormat(format)
            self.text_editor.mergeCurrentCharFormat(format)
    
    # Fixed bullet list insertion method using QTextListFormat        
    def insert_bullet_list(self):
        try:
            cursor = self.text_editor.textCursor()
            
            # Create a list format for bullet points
            list_format = QTextListFormat()
            list_format.setStyle(QTextListFormat.ListDisc)
            list_format.setIndent(1)
            
            # Check if we're already in a list
            current_list = cursor.currentList()
            
            if current_list:
                # We're in a list, so exit it
                cursor.beginEditBlock()
                cursor.insertBlock()
                cursor.endEditBlock()
            else:
                # We're not in a list, so create one
                cursor.beginEditBlock()
                cursor.createList(list_format)
                cursor.endEditBlock()
                
            self.text_editor.setFocus()
        except Exception as e:
            self.show_error(f"Error inserting bullet list: {str(e)}")
    
    # Fixed numbered list insertion method using QTextListFormat
    def insert_numbered_list(self):
        try:
            cursor = self.text_editor.textCursor()
            
            # Create a list format for numbered list
            list_format = QTextListFormat()
            list_format.setStyle(QTextListFormat.ListDecimal)
            list_format.setIndent(1)
            
            # Check if we're already in a list
            current_list = cursor.currentList()
            
            if current_list:
                # We're in a list, so exit it
                cursor.beginEditBlock()
                cursor.insertBlock()
                cursor.endEditBlock()
            else:
                # We're not in a list, so create one
                cursor.beginEditBlock()
                cursor.createList(list_format)
                cursor.endEditBlock()
                
            self.text_editor.setFocus()
        except Exception as e:
            self.show_error(f"Error inserting numbered list: {str(e)}")
        
    def add_attachment(self):
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Files to Attach")
            for file_path in file_paths:
                if file_path:
                    file_name = os.path.basename(file_path)
                    self.attachments.append(file_path)
                    self.composition_panel.attachment_panel.attachment_list.addItem(file_name)
        except Exception as e:
            self.show_error(f"Error adding attachment: {str(e)}")
    
    # Fixed attachment removal functionality            
    def remove_attachment(self):
        try:
            selected_items = self.composition_panel.attachment_panel.attachment_list.selectedItems()
            if not selected_items:
                return
                
            # Process items in reverse order to avoid index shifting issues
            rows_to_remove = []
            for item in selected_items:
                rows_to_remove.append(self.composition_panel.attachment_panel.attachment_list.row(item))
            
            # Sort in descending order to remove from bottom to top
            rows_to_remove.sort(reverse=True)
            
            for row in rows_to_remove:
                self.composition_panel.attachment_panel.attachment_list.takeItem(row)
                if row < len(self.attachments):
                    del self.attachments[row]
                    
        except Exception as e:
            self.show_error(f"Error removing attachment: {str(e)}")
            
    def validate_email(self):
        try:
            # Show validation in progress
            self.statusBar().showMessage("Validating email...")
            self.validation_panel.set_result("<p>Validating with Gemini...</p>")
            self.validation_panel.show_actions(False)
            self.validation_panel.show_refine_button(False)
            self.refined_panel.setVisible(False)
            
            # Get email content
            recipient = self.composition_panel.recipient_input.text()
            subject = self.composition_panel.subject_input.text()
            body = self.text_editor.toHtml()
            
            # Create validation prompt
            prompt = self.create_validation_prompt(recipient, subject, body, self.attachments)
            
            # Call Gemini API
            validation_result = self.call_gemini_api(prompt)
            
            # Display validation result
            self.validation_panel.set_result(validation_result)
            
            # Show appropriate buttons based on validation result
            if "not ok" in validation_result.lower():
                self.validation_panel.show_actions(True)
                self.validation_panel.show_refine_button(True)
                self.statusBar().showMessage("Validation failed. Please review the issues.")
            else:
                self.validation_panel.show_refine_button(True)
                self.statusBar().showMessage("Validation successful!")
        except Exception as e:
            self.show_error(f"Error during validation: {str(e)}")
    
    # Improved subject validation to be less nitpicky        
    def create_validation_prompt(self, recipient, subject, body, attachments):
        try:
            # Extract plain text from HTML for validation
            plain_body = self.text_editor.toPlainText()
            
            # Create a comprehensive prompt for Gemini with less strict subject validation
            prompt = f"""
            Please check if this email is technically correct and ready to send. 
            
            Respond with ONLY "yes" if everything is correct.
            
            If there are issues, respond with "not ok" followed by a numbered list of specific issues that need correction.
            
            Check for:
            1. Missing or invalid recipient email addresses (also check if the name of recipient in the email address is same as the name called in email body (if it exists))
            2. Empty subject line (don't be too strict about subject content, just ensure it conveys the overall meaning as the email body)
            3. Empty body or incomplete sentences
            4. Unclosed quotes, parentheses, or brackets
            5. Mentions of attachments without actual attachments being present
            6. Grammar or spelling issues that significantly impact understanding
            7. Any other technical problems that would significantly prevent effective communication
            
            Email details:
            TO: {recipient}
            SUBJECT: {subject}
            BODY: {plain_body}
            ATTACHMENTS: {', '.join([os.path.basename(a) for a in attachments]) if attachments else 'None'}
            
            Remember: Respond with ONLY "yes" if everything is correct. Otherwise, respond with "not ok" followed by numbered issues.
            """
            return prompt
        except Exception as e:
            self.show_error(f"Error creating validation prompt: {str(e)}")
            return "Error creating validation prompt"
        
    def call_gemini_api(self, prompt):
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={self.api_key}"
            headers = {'Content-Type': 'application/json'}
            data = {
                "contents": [{
                    "parts": [{"text": prompt}]
                }]
            }
            
            response = requests.post(url, headers=headers, json=data)
            response_json = response.json()
            
            if 'candidates' in response_json and len(response_json['candidates']) > 0:
                if 'content' in response_json['candidates'][0] and 'parts' in response_json['candidates'][0]['content']:
                    result = response_json['candidates'][0]['content']['parts'][0]['text']
                    # Format the result as HTML
                    result_html = "<p>" + result.replace("\n\n", "</p><p>").replace("\n", "<br>") + "</p>"
                    return result_html
            
            return "<p>Error: Unable to get a valid response from Gemini.</p>"
        except Exception as e:
            return f"<p>Error calling Gemini API: {str(e)}</p>"
    
    def refine_email(self):
        try:
            # Show refinement in progress
            self.statusBar().showMessage("Refining email...")
            self.refined_panel.setVisible(True)
            self.refined_panel.set_content("<p>Refining with Gemini...</p>")
            
            # Get email content
            recipient = self.composition_panel.recipient_input.text()
            subject = self.composition_panel.subject_input.text()
            body_text = self.text_editor.toPlainText()
            body_html = self.text_editor.toHtml()
            
            # Get validation result if available
            validation_result = self.validation_panel.result_area.toPlainText()
            
            # Create refinement prompt based on validation result
            if "not ok" in validation_result.lower():
                prompt = self.create_full_refinement_prompt(recipient, subject, body_text, validation_result)
            else:
                prompt = self.create_minimal_refinement_prompt(recipient, subject, body_text)
            
            # Call Gemini API
            refined_content = self.call_gemini_api(prompt)
            
            # Parse and display refined content
            self.parse_refined_content(refined_content, body_html)
            
            # Display refined content
            refined_html = f"""
            <h3>Refined Email</h3>
            <p><strong>Subject:</strong> {self.refined_subject}</p>
            <div style="border-top: 1px solid #cccccc; margin: 10px 0;"></div>
            {self.refined_body_html}
            """
            self.refined_panel.set_content(refined_html)
            self.statusBar().showMessage("Email refined successfully!")
            
        except Exception as e:
            self.show_error(f"Error refining email: {str(e)}")
            self.refined_panel.set_content(f"<p>Error refining email: {str(e)}</p>")
    
    def create_minimal_refinement_prompt(self, recipient, subject, body):
        return f"""
        There are no major errors in this email, but please make minimal improvements to enhance clarity, professionalism, and effectiveness.
        
        Original email:
        TO: {recipient}
        SUBJECT: {subject}
        BODY:
        {body}
        
        Please provide the refined version in this exact format:
        SUBJECT: [refined subject]
        BODY:
        [refined body]
        
        Keep the same meaning and tone—just make small improvements to grammar, clarity, and professionalism.
        """

    def create_full_refinement_prompt(self, recipient, subject, body, validation_result):
        return f"""
        Please refine this email to fix all issues and improve its clarity, professionalism, and effectiveness.
        
        Original email:
        TO: {recipient}
        SUBJECT: {subject}
        BODY:
        {body}
        
        Validation feedback:
        {validation_result}
        
        Please provide the refined version in this exact format:
        SUBJECT: [refined subject]
        BODY:
        [refined body]
        
        Fix all issues mentioned in the validation feedback and make any other improvements needed. Don't add unnessary information by yourself. Just refine where needed.
        """
    
    def parse_refined_content(self, refined_content, original_html):
        try:
            # First, get the plain text version by removing any HTML tags
            plain_refined_content = refined_content
            # Remove HTML paragraph tags and breaks
            plain_refined_content = plain_refined_content.replace("<p>", "").replace("</p>", "\n\n")
            plain_refined_content = plain_refined_content.replace("<br>", "\n")
            
            # Extract subject and body from refined content
            if "SUBJECT:" in plain_refined_content and "BODY:" in plain_refined_content:
                # Extract subject
                subject_start = plain_refined_content.find("SUBJECT:") + 8
                subject_end = plain_refined_content.find("BODY:")
                self.refined_subject = plain_refined_content[subject_start:subject_end].strip()
                
                # Extract body
                body_start = plain_refined_content.find("BODY:") + 5
                self.refined_body_text = plain_refined_content[body_start:].strip()
                
                # Convert to HTML while preserving formatting
                self.refined_body_html = self.text_to_html(self.refined_body_text, original_html)
            else:
                # Fallback if format is not as expected
                self.refined_subject = "Refined Subject"
                self.refined_body_html = f"<p>{refined_content}</p>"
                self.refined_body_text = refined_content
        except Exception as e:
            self.show_error(f"Error parsing refined content: {str(e)}")
            self.refined_subject = "Error in refinement"
            self.refined_body_html = f"<p>Error parsing refined content: {str(e)}</p>"
            self.refined_body_text = f"Error parsing refined content: {str(e)}"

    def text_to_html(self, text, original_html):
        """Convert plain text to HTML while trying to preserve formatting from original HTML"""
        # Basic conversion of plain text to HTML
        html = ""
        paragraphs = text.split('\n\n')
        
        for paragraph in paragraphs:
            if paragraph.strip():
                if paragraph.strip().startswith('- '):
                    # Convert to unordered list
                    items = paragraph.split('\n- ')
                    html += "<ul>"
                    for item in items:
                        if item.strip():
                            html += f"<li>{item.strip()}</li>"
                    html += "</ul>"
                elif any(line.strip() and line.strip()[0].isdigit() and line.strip()[1:].startswith('. ') for line in paragraph.split('\n')):
                    # Convert to ordered list
                    items = paragraph.split('\n')
                    html += "<ol>"
                    for item in items:
                        if item.strip() and item.strip()[0].isdigit() and item.strip()[1:].startswith('. '):
                            content = item.strip()[item.strip().find('.')+1:].strip()
                            html += f"<li>{content}</li>"
                    html += "</ol>"
                else:
                    # Regular paragraph
                    lines = paragraph.split('\n')
                    html += f"<p>{'<br>'.join(lines)}</p>"
        
        return html
    
    def insert_refined_content(self):
        try:
            # Insert refined subject and body into the form
            self.composition_panel.subject_input.setText(self.refined_subject)
            
            # Clear current content and insert refined HTML
            self.text_editor.clear()
            self.text_editor.setHtml(self.refined_body_html)
            
            # Hide refined content panel
            self.refined_panel.setVisible(False)
            
            self.statusBar().showMessage("Refined content inserted successfully!")
        except Exception as e:
            self.show_error(f"Error inserting refined content: {str(e)}")
            
    def confirm_send(self):
        try:
            # First validate
            self.validate_email()
            
            # If validation passes, ask for confirmation
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("Confirm Send")
            msg_box.setText("Are you ready to send this email?")
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg_box.setDefaultButton(QMessageBox.No)
            
            if msg_box.exec_() == QMessageBox.Yes:
                self.send_email()
        except Exception as e:
            self.show_error(f"Error confirming send: {str(e)}")
            
    def send_email(self):
        try:
            # Show sending in progress
            self.statusBar().showMessage("Sending email...")
            
            # Get email content
            recipient = self.composition_panel.recipient_input.text()
            subject = self.composition_panel.subject_input.text()
            body_html = self.text_editor.toHtml()
            
            # Create message
            msg = MIMEMultipart('alternative')
            msg['From'] = self.email
            msg['To'] = recipient
            msg['Subject'] = subject
            
            # Attach HTML body
            msg.attach(MIMEText(body_html, 'html'))
            
            # Attach files
            for file_path in self.attachments:
                with open(file_path, 'rb') as file:
                    attachment = MIMEApplication(file.read())
                    file_name = os.path.basename(file_path)
                    attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
                    msg.attach(attachment)
            
            # Connect to SMTP server and send
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(self.email, self.password)
                server.send_message(msg)
            
            self.show_success("Email sent successfully!")
            self.clear_form()
            
        except Exception as e:
            self.show_error(f"Failed to send email: {str(e)}")
            
    def clear_form(self):
        try:
            self.composition_panel.recipient_input.clear()
            self.composition_panel.subject_input.clear()
            self.text_editor.clear()
            self.composition_panel.attachment_panel.attachment_list.clear()
            self.attachments.clear()
            self.validation_panel.result_area.clear()
            self.validation_panel.show_actions(False)
            self.validation_panel.show_refine_button(False)
            self.refined_panel.setVisible(False)
            self.statusBar().showMessage("Form cleared")
        except Exception as e:
            self.show_error(f"Error clearing form: {str(e)}")
    
    def show_error(self, message):
        """Display error message with consistent styling"""
        error_box = QMessageBox(self)
        error_box.setWindowTitle("Error")
        error_box.setText(message)
        error_box.setIcon(QMessageBox.Critical)
        error_box.setStandardButtons(QMessageBox.Ok)
        error_box.exec_()
        self.statusBar().showMessage(f"Error: {message}")
        
    def show_success(self, message):
        """Display success message with consistent styling"""
        success_box = QMessageBox(self)
        success_box.setWindowTitle("Success")
        success_box.setText(message)
        success_box.setIcon(QMessageBox.Information)
        success_box.setStandardButtons(QMessageBox.Ok)
        success_box.exec_()
        self.statusBar().showMessage(message)
        
    def keyPressEvent(self, event):
        # Handle key press events
        if event.key() == Qt.Key_E:
            # Just continue editing (default behavior)
            pass
        elif event.key() == Qt.Key_A:
            # Abort
            self.clear_form()
        elif event.key() == Qt.Key_Return and (event.modifiers() & Qt.ControlModifier):
            # Ctrl+Enter to send
            self.confirm_send()
        else:
            super().keyPressEvent(event)

def main():
    try:
        app = QApplication(sys.argv)
        
        # Set application font
        app_font = QFont("Segoe UI", 10)
        app.setFont(app_font)
        
        window = EmailComposer()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Application error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
