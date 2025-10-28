import tkinter as tk
from tkinter import ttk, messagebox
from lessonnotegeneratorupdated import LessonNoteGenerator
from examgeneratorupdated import ExamQuestionGenerator
import os


class ApplicationLauncher:
    """Main application launcher that provides access to both tools."""
    
    # Modern color scheme
    PRIMARY_COLOR = "#4a6da7"
    SECONDARY_COLOR = "#6dd5fa"
    BG_COLOR = "#f8f9fa"
    TEXT_COLOR = "#2d3748"
    
    # Fonts
    TITLE_FONT = ("Segoe UI", 24, "bold")
    BUTTON_FONT = ("Segoe UI", 12, "bold")
    
    def __init__(self, root):
        self.root = root
        self.root.title("Avalon Educational Tools")
        self.root.geometry("600x400")
        self.root.configure(bg=self.BG_COLOR)
        
                
        self._setup_ui()
    
    def _setup_ui(self):
        """Create the launcher interface."""
        # Header frame
        header_frame = tk.Frame(self.root, bg=self.BG_COLOR, padx=20, pady=30)
        header_frame.pack(fill="x")
        
        title_label = tk.Label(
            header_frame,
            text="Avalon Educational Tools",
            font=self.TITLE_FONT,
            bg=self.BG_COLOR,
            fg=self.PRIMARY_COLOR,
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            header_frame,
            text="Choose a tool to get started",
            font=("Segoe UI", 12),
            bg=self.BG_COLOR,
            fg="#718096",
        )
        subtitle_label.pack(pady=(10, 30))
        
        # Button container
        button_frame = tk.Frame(self.root, bg=self.BG_COLOR, padx=50, pady=20)
        button_frame.pack(expand=True, fill="both")
        
        # Lesson Note Generator button
        lesson_btn = tk.Button(
            button_frame,
            text="Lesson Note Generator",
            font=self.BUTTON_FONT,
            bg=self.PRIMARY_COLOR,
            fg="white",
            activebackground="#3a5a8f",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=15,
            command=self.launch_lesson_note_generator,
        )
        lesson_btn.pack(fill="x", pady=10, ipadx=10, ipady=5)
        
        # Exam Question Generator button
        exam_btn = tk.Button(
            button_frame,
            text="Exam Question Generator",
            font=self.BUTTON_FONT,
            bg=self.PRIMARY_COLOR,
            fg="white",
            activebackground="#3a5a8f",
            activeforeground="white",
            relief="flat",
            padx=20,
            pady=15,
            command=self.launch_exam_generator,
        )
        exam_btn.pack(fill="x", pady=10, ipadx=10, ipady=5)
        
        # Footer
        footer_frame = tk.Frame(self.root, bg=self.BG_COLOR, padx=20, pady=20)
        footer_frame.pack(fill="x", side="bottom")
        
        footer_label = tk.Label(
            footer_frame,
            text="© Avalon Educational Tools",
            font=("Segoe UI", 9),
            bg=self.BG_COLOR,
            fg="#718096",
        )
        footer_label.pack()
    
    def launch_lesson_note_generator(self):
        """Launch the Lesson Note Generator application."""
        self.root.withdraw()  # Hide the launcher window
        lesson_window = tk.Toplevel()
        LessonNoteGenerator(lesson_window)
        lesson_window.protocol("WM_DELETE_WINDOW", lambda: self.on_child_close(lesson_window))
    
    def launch_exam_generator(self):
        """Launch the Exam Question Generator application."""
        self.root.withdraw()  # Hide the launcher window
        exam_window = tk.Toplevel()
        ExamQuestionGenerator(exam_window)
        exam_window.protocol("WM_DELETE_WINDOW", lambda: self.on_child_close(exam_window))
    
    def on_child_close(self, window):
        """Handle child window closing."""
        window.destroy()
        self.root.deiconify()  # Show the launcher window again


def main():
    """Entry point for the application."""
    root = tk.Tk()
    
    # Set Windows 10/11 theme if available
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    
    app = ApplicationLauncher(root)
    root.mainloop()


if __name__ == "__main__":
    main()