import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import smtplib
import os
import re
import threading
from email.message import EmailMessage
from pathlib import Path
import mimetypes

class EmailSenderPro:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Email Sender")
        
        # Full screen
        self.root.state('zoomed')  # Windows
        self.root.minsize(900, 700)
        
        # Color scheme
        self.colors = {
            'primary': '#2563eb',      # Blue
            'secondary': '#10b981',    # Green
            'danger': '#ef4444',       # Red
            'bg': '#f8fafc',          # Light gray
            'card': '#ffffff',        # White
            'text': '#1e293b',        # Dark gray
            'border': '#e2e8f0'       # Border gray
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Variables
        self.uploaded_file_path = None
        self.uploaded_df = None
        self.email_column = None
        self.attachments = []
        self.stop_sending = False
        
        # Style configuration
        self.setup_styles()
        
        # Create UI
        self.create_header()
        self.create_tabs()
        
        # Center window
        self.center_window()
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure notebook (tabs)
        style.configure('TNotebook', background=self.colors['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', 
                       padding=[20, 10],
                       font=('Segoe UI', 10, 'bold'),
                       background=self.colors['card'])
        style.map('TNotebook.Tab',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', 'white')])
        
        # Configure frames
        style.configure('Card.TFrame', background=self.colors['card'], relief='flat')
        style.configure('TFrame', background=self.colors['bg'])
        
        # Configure labels
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['card'],
                       foreground=self.colors['text'])
        style.configure('TLabel',
                       font=('Segoe UI', 10),
                       background=self.colors['card'],
                       foreground=self.colors['text'])
    
    def center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def create_header(self):
        header = tk.Frame(self.root, bg=self.colors['primary'], height=80)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        
        title = tk.Label(header, 
                        text="📧 Professional Email Sender",
                        font=('Segoe UI', 20, 'bold'),
                        bg=self.colors['primary'],
                        fg='white')
        title.pack(pady=20)
    
    def create_tabs(self):
        # Create notebook
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.bulk_tab = ttk.Frame(self.notebook)
        self.individual_tab = ttk.Frame(self.notebook)
        self.settings_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.bulk_tab, text='  📤 Bulk Email  ')
        self.notebook.add(self.individual_tab, text='  ✉️ Individual Email  ')
        self.notebook.add(self.settings_tab, text='  🔐 Gmail Settings  ')
        
        # Build tab contents
        self.build_bulk_tab()
        self.build_individual_tab()
        self.build_settings_tab()
    
    def build_bulk_tab(self):
        # Two-column layout
        main_container = tk.Frame(self.bulk_tab, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Left Column - File Upload & Preview
        left_col = tk.Frame(main_container, bg=self.colors['bg'])
        left_col.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # File Upload Card
        file_card = self.create_card(left_col, "📁 Upload Email List")
        file_card.pack(fill='x', pady=(0, 15))
        
        upload_frame = tk.Frame(file_card, bg=self.colors['card'])
        upload_frame.pack(fill='x', padx=20, pady=15)
        
        self.file_label = tk.Label(upload_frame, text="No file selected",
                                   font=('Segoe UI', 10), bg=self.colors['card'], fg='gray')
        self.file_label.pack(side='left', fill='x', expand=True)
        
        self.create_button(upload_frame, "📂 Choose File", self.upload_file, 
                          self.colors['primary']).pack(side='right')
        
        # Column Selection
        col_frame = tk.Frame(file_card, bg=self.colors['card'])
        col_frame.pack(fill='x', padx=20, pady=(0, 15))
        
        tk.Label(col_frame, text="Email Column:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(side='left', padx=(0, 10))
        
        self.column_combo = ttk.Combobox(col_frame, state='disabled', 
                                        font=('Segoe UI', 10), width=30)
        self.column_combo.pack(side='left', fill='x', expand=True)
        self.column_combo.bind('<<ComboboxSelected>>', self.on_column_selected)
        
        # Preview Card
        preview_card = self.create_card(left_col, "👁️ Email Preview")
        preview_card.pack(fill='both', expand=True, pady=(0, 15))
        
        self.preview_text = self.create_scrolled_text(preview_card, height=15)
        self.preview_text.insert('1.0', 'Upload a file to see email addresses...')
        self.preview_text.config(state='disabled')
        
        # Send Button
        self.create_button(left_col, "🚀 Send Bulk Emails", self.send_bulk_emails,
                          self.colors['secondary'], large=True).pack()
        
        # Right Column - Compose Email
        right_col = tk.Frame(main_container, bg=self.colors['bg'])
        right_col.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        compose_card = self.create_card(right_col, "✉️ Compose Email")
        compose_card.pack(fill='both', expand=True, pady=(0, 15))
        
        compose_frame = tk.Frame(compose_card, bg=self.colors['card'])
        compose_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Subject
        tk.Label(compose_frame, text="Subject:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(anchor='w', pady=(0, 5))
        self.bulk_subject = self.create_entry(compose_frame)
        self.bulk_subject.pack(fill='x', pady=(0, 15))
        
        # Body
        tk.Label(compose_frame, text="Message Body:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(anchor='w', pady=(0, 5))
        
        self.create_format_toolbar(compose_frame, lambda t: self.toggle_format_bulk(t))
        
        self.bulk_body = self.create_scrolled_text(compose_frame, height=12)
        self.bulk_body.tag_configure('bold', font=('Segoe UI', 10, 'bold'))
        self.bulk_body.tag_configure('italic', font=('Segoe UI', 10, 'italic'))
        
        # Attachments
        attach_card = self.create_card(right_col, "📎 Attachments")
        attach_card.pack(fill='x')
        
        attach_frame = tk.Frame(attach_card, bg=self.colors['card'])
        attach_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        btn_frame = tk.Frame(attach_frame, bg=self.colors['card'])
        btn_frame.pack(fill='x', pady=(0, 10))
        
        self.create_button(btn_frame, "📎 Add", self.add_attachment, 
                          self.colors['primary'], small=True).pack(side='left')
        self.create_button(btn_frame, "🗑️ Clear All", self.clear_attachments,
                          self.colors['danger'], small=True).pack(side='left', padx=10)
        
        self.bulk_attachment_frame = tk.Frame(attach_frame, bg=self.colors['card'])
        self.bulk_attachment_frame.pack(fill='both', expand=True)
        
        # Show initial empty state
        tk.Label(self.bulk_attachment_frame,
                text="No attachments added",
                font=('Segoe UI', 9, 'italic'),
                bg=self.colors['card'],
                fg='gray').pack(pady=10)
    
    def build_individual_tab(self):
        # Two-column layout
        main_container = tk.Frame(self.individual_tab, bg=self.colors['bg'])
        main_container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Left Column - Recipient
        left_col = tk.Frame(main_container, bg=self.colors['bg'])
        left_col.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        recipient_card = self.create_card(left_col, "👤 Recipient")
        recipient_card.pack(fill='x', pady=(0, 15))
        
        recipient_frame = tk.Frame(recipient_card, bg=self.colors['card'])
        recipient_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(recipient_frame, text="Email Address:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(anchor='w', pady=(0, 5))
        
        self.individual_email = self.create_entry(recipient_frame)
        self.individual_email.pack(fill='x', pady=(0, 5))
        
        # Info
        info_card = self.create_card(left_col, "ℹ️ Instructions")
        info_card.pack(fill='both', expand=True, pady=(0, 15))
        
        info_frame = tk.Frame(info_card, bg=self.colors['card'])
        info_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        instructions = """1. Enter recipient's email address
2. Compose your message on the right
3. Add attachments if needed
4. Click Send Email button

⚠️ Make sure Gmail Settings are configured!"""
        
        tk.Label(info_frame, text=instructions, font=('Segoe UI', 10),
                bg=self.colors['card'], fg='#64748b', justify='left',
                anchor='w').pack(fill='both')
        
        # Send Button
        self.create_button(left_col, "✉️ Send Email", self.send_individual_email,
                          self.colors['secondary'], large=True).pack()
        
        # Right Column - Compose
        right_col = tk.Frame(main_container, bg=self.colors['bg'])
        right_col.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        compose_card = self.create_card(right_col, "✉️ Compose Email")
        compose_card.pack(fill='both', expand=True, pady=(0, 15))
        
        compose_frame = tk.Frame(compose_card, bg=self.colors['card'])
        compose_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        # Subject
        tk.Label(compose_frame, text="Subject:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(anchor='w', pady=(0, 5))
        self.individual_subject = self.create_entry(compose_frame)
        self.individual_subject.pack(fill='x', pady=(0, 15))
        
        # Body
        tk.Label(compose_frame, text="Message Body:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['card']).pack(anchor='w', pady=(0, 5))
        
        self.create_format_toolbar(compose_frame, lambda t: self.toggle_format_individual(t))
        
        self.individual_body = self.create_scrolled_text(compose_frame, height=12)
        self.individual_body.tag_configure('bold', font=('Segoe UI', 10, 'bold'))
        self.individual_body.tag_configure('italic', font=('Segoe UI', 10, 'italic'))
        
        # Attachments
        attach_card = self.create_card(right_col, "📎 Attachments")
        attach_card.pack(fill='x')
        
        attach_frame = tk.Frame(attach_card, bg=self.colors['card'])
        attach_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        btn_frame = tk.Frame(attach_frame, bg=self.colors['card'])
        btn_frame.pack(fill='x', pady=(0, 10))
        
        self.create_button(btn_frame, "📎 Add", self.add_attachment,
                          self.colors['primary'], small=True).pack(side='left')
        self.create_button(btn_frame, "🗑️ Clear All", self.clear_attachments,
                          self.colors['danger'], small=True).pack(side='left', padx=10)
        
        self.individual_attachment_frame = tk.Frame(attach_frame, bg=self.colors['card'])
        self.individual_attachment_frame.pack(fill='both', expand=True)
        
        # Show initial empty state
        tk.Label(self.individual_attachment_frame,
                text="No attachments added",
                font=('Segoe UI', 9, 'italic'),
                bg=self.colors['card'],
                fg='gray').pack(pady=10)
    
    def build_settings_tab(self):
        container = tk.Frame(self.settings_tab, bg=self.colors['bg'])
        container.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Gmail Credentials
        cred_card = self.create_card(container, "🔐 Gmail Account Settings")
        cred_card.pack(fill='x', pady=(0, 15))
        
        cred_frame = tk.Frame(cred_card, bg=self.colors['card'])
        cred_frame.pack(fill='x', padx=20, pady=15)
        
        tk.Label(cred_frame, text="Gmail Address:", font=('Segoe UI', 10, 'bold'), bg=self.colors['card']).grid(row=0, column=0, sticky='w', pady=12)
        self.gmail_email = tk.Entry(cred_frame, font=('Segoe UI', 11), relief='flat', bg='#f1f5f9')
        self.gmail_email.grid(row=0, column=1, sticky='ew', pady=12, padx=10)
        
        tk.Label(cred_frame, text="App Password:", font=('Segoe UI', 10, 'bold'), bg=self.colors['card']).grid(row=1, column=0, sticky='w', pady=12)
        self.gmail_password = tk.Entry(cred_frame, font=('Segoe UI', 11), show='•', relief='flat', bg='#f1f5f9')
        self.gmail_password.grid(row=1, column=1, sticky='ew', pady=12, padx=10)
        
        cred_frame.columnconfigure(1, weight=1)
        
        # Instructions Card
        info_card = self.create_card(container, "ℹ️ How to Get Gmail App Password")
        info_card.pack(fill='x', pady=(0, 15))
        
        info_frame = tk.Frame(info_card, bg=self.colors['card'])
        info_frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        instructions = """1. Go to Google Account Security settings
2. Enable 2-Step Verification if not already enabled
3. Search for "App passwords"
4. Create a new app password for "Mail"
5. Copy the 16-character password
6. Paste it in the App Password field above

⚠️ Use the App Password, NOT your regular Gmail password!"""
        
        tk.Label(info_frame,
                text=instructions,
                font=('Segoe UI', 9),
                bg=self.colors['card'],
                fg='#64748b',
                justify='left',
                anchor='w').pack(fill='both')
    
    def create_button(self, parent, text, command, bg_color, small=False, large=False):
        """Create a styled button component"""
        if large:
            font_size, px, py = 12, 30, 12
        elif small:
            font_size, px, py = 9, 15, 6
        else:
            font_size, px, py = 10, 20, 8
        
        return tk.Button(parent, text=text, command=command,
                        font=('Segoe UI', font_size, 'bold'),
                        bg=bg_color, fg='white', cursor='hand2',
                        relief='flat', padx=px, pady=py)
    
    def create_entry(self, parent):
        """Create a styled entry component"""
        return tk.Entry(parent, font=('Segoe UI', 10), relief='flat', bg='#f1f5f9')
    
    def create_scrolled_text(self, parent, height=10):
        """Create a styled scrolled text widget"""
        frame = tk.Frame(parent, bg=self.colors['card'])
        frame.pack(fill='both', expand=True, padx=20, pady=15)
        
        text = scrolledtext.ScrolledText(frame, height=height,
                                        font=('Segoe UI', 10),
                                        relief='flat', bg='#f1f5f9',
                                        padx=10, pady=10, wrap='word')
        text.pack(fill='both', expand=True)
        return text
    
    def create_format_toolbar(self, parent, toggle_callback):
        """Create formatting toolbar component"""
        toolbar = tk.Frame(parent, bg=self.colors['card'])
        toolbar.pack(fill='x', pady=(0, 5))
        
        tk.Button(toolbar, text="B", font=('Segoe UI', 9, 'bold'),
                 command=lambda: toggle_callback('bold'),
                 width=3, relief='flat', bg='#e2e8f0', cursor='hand2').pack(side='left', padx=2)
        tk.Button(toolbar, text="I", font=('Segoe UI', 9, 'italic'),
                 command=lambda: toggle_callback('italic'),
                 width=3, relief='flat', bg='#e2e8f0', cursor='hand2').pack(side='left', padx=2)
        
        tk.Label(toolbar, text="  Ctrl+B: Bold  |  Ctrl+I: Italic",
                font=('Segoe UI', 8), bg=self.colors['card'], fg='gray').pack(side='left', padx=10)
    
    def create_card(self, parent, title):
        """Create a card component with title"""
        card = tk.Frame(parent, bg=self.colors['card'], relief='flat', borderwidth=1)
        card.configure(highlightbackground=self.colors['border'], highlightthickness=1)
        
        header = tk.Frame(card, bg=self.colors['card'])
        header.pack(fill='x', padx=15, pady=12)
        
        tk.Label(header, text=title, font=('Segoe UI', 12, 'bold'),
                bg=self.colors['card'], fg=self.colors['text']).pack(anchor='w')
        
        return card
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Email List File",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Read file
            if file_path.endswith('.csv'):
                self.uploaded_df = pd.read_csv(file_path)
            else:
                self.uploaded_df = pd.read_excel(file_path)
            
            self.uploaded_file_path = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"✅ {filename}", fg=self.colors['secondary'])
            
            # Populate column dropdown
            columns = list(self.uploaded_df.columns)
            self.column_combo['values'] = columns
            self.column_combo['state'] = 'readonly'
            
            # Auto-select if there's an 'email' column
            email_cols = [col for col in columns if 'email' in col.lower()]
            if email_cols:
                self.column_combo.set(email_cols[0])
                self.on_column_selected()
            else:
                self.column_combo.set('')
            
            messagebox.showinfo("Success", f"File loaded successfully!\n{len(self.uploaded_df)} rows found.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
    
    def on_column_selected(self, event=None):
        if not self.column_combo.get() or self.uploaded_df is None:
            return
        
        self.email_column = self.column_combo.get()
        emails = self.uploaded_df[self.email_column].dropna().unique()
        
        # Update preview
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', 'end')
        self.preview_text.insert('1.0', f"Found {len(emails)} unique email(s):\n\n")
        
        for i, email in enumerate(emails[:50], 1):  # Show first 50
            self.preview_text.insert('end', f"{i}. {email}\n")
        
        if len(emails) > 50:
            self.preview_text.insert('end', f"\n... and {len(emails) - 50} more")
        
        self.preview_text.config(state='disabled')
    
    def add_attachment(self):
        files = filedialog.askopenfilenames(
            title="Select Attachments",
            filetypes=[
                ("All files", "*.*"),
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx *.doc"),
                ("Images", "*.png *.jpg *.jpeg *.gif"),
                ("Text files", "*.txt")
            ]
        )
        
        if files:  # Only update if files were selected
            for file in files:
                if file not in self.attachments:
                    self.attachments.append(file)
            
            print(f"DEBUG: Added {len(files)} files. Total attachments: {len(self.attachments)}")
            print(f"DEBUG: Attachments list: {self.attachments}")
            self.update_attachment_list()
    
    def remove_attachment(self, file_path):
        if file_path in self.attachments:
            self.attachments.remove(file_path)
        self.update_attachment_list()
    
    def clear_attachments(self):
        if self.attachments and messagebox.askyesno("Confirm", "Remove all attachments?"):
            self.attachments.clear()
            self.update_attachment_list()
    
    def update_attachment_list(self):
        # Update both bulk and individual attachment frames
        print(f"DEBUG: update_attachment_list called. Attachments count: {len(self.attachments)}")
        for frame_attr in ['bulk_attachment_frame', 'individual_attachment_frame']:
            print(f"DEBUG: Checking frame: {frame_attr}, exists: {hasattr(self, frame_attr)}")
            if hasattr(self, frame_attr):
                frame = getattr(self, frame_attr)
                print(f"DEBUG: Updating frame {frame_attr}")
                # Clear frame
                for widget in frame.winfo_children():
                    widget.destroy()
                
                if not self.attachments:
                    tk.Label(frame,
                            text="No attachments added",
                            font=('Segoe UI', 9, 'italic'),
                            bg=self.colors['card'],
                            fg='gray').pack(pady=10)
                else:
                    # Create scrollable frame for attachments
                    canvas = tk.Canvas(frame, bg=self.colors['card'], 
                                      height=min(150, len(self.attachments) * 35),
                                      highlightthickness=0)
                    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
                    scroll_frame = tk.Frame(canvas, bg=self.colors['card'])
                    
                    scroll_frame.bind("<Configure>",
                                     lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                    
                    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                    canvas.configure(yscrollcommand=scrollbar.set)
                    
                    canvas.pack(side="left", fill="both", expand=True)
                    scrollbar.pack(side="right", fill="y")
                    
                    for file_path in self.attachments:
                        file_frame = tk.Frame(scroll_frame, bg='#f1f5f9', 
                                             highlightbackground='#cbd5e1',
                                             highlightthickness=1)
                        file_frame.pack(fill='x', pady=3, padx=5)
                        
                        filename = os.path.basename(file_path)
                        # Truncate long filenames
                        display_name = filename if len(filename) <= 40 else filename[:37] + '...'
                        
                        tk.Label(file_frame,
                                text=f"📄 {display_name}",
                                font=('Segoe UI', 9),
                                bg='#f1f5f9',
                                anchor='w').pack(side='left', fill='x', expand=True, padx=10, pady=6)
                        
                        tk.Button(file_frame,
                                 text="✖",
                                 command=lambda f=file_path: self.remove_attachment(f),
                                 font=('Segoe UI', 9, 'bold'),
                                 bg=self.colors['danger'],
                                 fg='white',
                                 cursor='hand2',
                                 relief='flat',
                                 width=3,
                                 height=1).pack(side='right', padx=5, pady=3)
    
    def toggle_format_bulk(self, tag):
        try:
            if self.bulk_body.tag_ranges('sel'):
                current_tags = self.bulk_body.tag_names('sel.first')
                if tag in current_tags:
                    self.bulk_body.tag_remove(tag, 'sel.first', 'sel.last')
                else:
                    self.bulk_body.tag_add(tag, 'sel.first', 'sel.last')
        except tk.TclError:
            pass
    
    def toggle_format_individual(self, tag):
        try:
            if self.individual_body.tag_ranges('sel'):
                current_tags = self.individual_body.tag_names('sel.first')
                if tag in current_tags:
                    self.individual_body.tag_remove(tag, 'sel.first', 'sel.last')
                else:
                    self.individual_body.tag_add(tag, 'sel.first', 'sel.last')
        except tk.TclError:
            pass
    
    def get_html_body(self, text_widget):
        """Convert text widget content with formatting to HTML"""
        html = ""
        index = "1.0"
        
        while True:
            # Get all tags at current position
            next_tag = None
            min_idx = None
            
            for tag in ('bold', 'italic'):
                ranges = text_widget.tag_nextrange(tag, index)
                if ranges:
                    if min_idx is None or text_widget.compare(ranges[0], '<', min_idx):
                        min_idx = ranges[0]
                        next_tag = (ranges[0], ranges[1], tag)
            
            if next_tag is None:
                # No more tags, get remaining text
                html += text_widget.get(index, 'end-1c')
                break
            
            # Add text before tag
            html += text_widget.get(index, next_tag[0])
            
            # Add tagged text
            tag_text = text_widget.get(next_tag[0], next_tag[1])
            if next_tag[2] == 'bold':
                html += f'<b>{tag_text}</b>'
            elif next_tag[2] == 'italic':
                html += f'<i>{tag_text}</i>'
            
            index = next_tag[1]
        
        # Convert newlines to <br>
        html = html.replace('\n', '<br>')
        return html
    
    def validate_credentials(self):
        """Validate Gmail credentials only"""
        email = self.gmail_email.get().strip()
        password = self.gmail_password.get().strip()
        
        if not email:
            messagebox.showwarning("Missing Info", "Please enter your Gmail address in Gmail Settings tab.")
            self.notebook.select(2)  # Switch to settings tab
            return False
        
        if not password:
            messagebox.showwarning("Missing Info", "Please enter your Gmail App Password in Gmail Settings tab.")
            self.notebook.select(2)
            return False
        
        return True
    
    def validate_bulk_email(self):
        """Validate bulk email fields"""
        subject = self.bulk_subject.get().strip()
        body = self.bulk_body.get('1.0', 'end-1c').strip()
        
        if not subject:
            messagebox.showwarning("Missing Info", "Please enter an email subject in the Bulk Email tab.")
            self.notebook.select(0)
            return False
        
        if not body:
            messagebox.showwarning("Missing Info", "Please enter email body text in the Bulk Email tab.")
            self.notebook.select(0)
            return False
        
        return True
    
    def validate_individual_email(self):
        """Validate individual email fields"""
        subject = self.individual_subject.get().strip()
        body = self.individual_body.get('1.0', 'end-1c').strip()
        
        if not subject:
            messagebox.showwarning("Missing Info", "Please enter an email subject in the Individual Email tab.")
            self.notebook.select(1)
            return False
        
        if not body:
            messagebox.showwarning("Missing Info", "Please enter email body text in the Individual Email tab.")
            self.notebook.select(1)
            return False
        
        return True
    
    def send_bulk_emails(self):
        """Send emails to all addresses in uploaded file"""
        if self.uploaded_df is None:
            messagebox.showwarning("No File", "Please upload a file with email addresses first.")
            return
        
        if not self.email_column:
            messagebox.showwarning("No Column", "Please select the email column.")
            return
        
        if not self.validate_credentials() or not self.validate_bulk_email():
            return
        
        # Get unique emails
        emails = self.uploaded_df[self.email_column].dropna().unique()
        
        if len(emails) == 0:
            messagebox.showwarning("No Emails", "No email addresses found in selected column.")
            return
        
        # Confirm
        if not messagebox.askyesno("Confirm", 
                                   f"Send email to {len(emails)} recipient(s)?"):
            return
        
        # Open progress window
        self.open_progress_window(list(emails), 'bulk')
    
    def send_individual_email(self):
        """Send email to manually entered address"""
        email = self.individual_email.get().strip()
        
        if not email:
            messagebox.showwarning("No Email", "Please enter a recipient email address.")
            return
        
        # Basic email validation
        if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
            messagebox.showwarning("Invalid Email", "Please enter a valid email address.")
            return
        
        if not self.validate_credentials() or not self.validate_individual_email():
            return
        
        # Confirm
        if not messagebox.askyesno("Confirm", f"Send email to {email}?"):
            return
        
        # Open progress window
        self.open_progress_window([email], 'individual')
    
    def open_progress_window(self, email_list, mode):
        """Open progress window and start sending"""
        self.stop_sending = False
        
        # Create progress window
        progress_win = tk.Toplevel(self.root)
        progress_win.title("Sending Emails")
        progress_win.geometry("600x500")
        progress_win.transient(self.root)
        progress_win.grab_set()
        
        # Center window
        progress_win.update_idletasks()
        x = (progress_win.winfo_screenwidth() // 2) - 300
        y = (progress_win.winfo_screenheight() // 2) - 250
        progress_win.geometry(f"600x500+{x}+{y}")
        
        # Header
        header = tk.Frame(progress_win, bg=self.colors['primary'])
        header.pack(fill='x')
        tk.Label(header,
                text="📨 Sending Emails...",
                font=('Segoe UI', 14, 'bold'),
                bg=self.colors['primary'],
                fg='white').pack(pady=15)
        
        # Progress info
        info_frame = tk.Frame(progress_win, bg='white')
        info_frame.pack(fill='x', padx=20, pady=10)
        
        progress_label = tk.Label(info_frame,
                                 text=f"Total: {len(email_list)} | Sent: 0 | Failed: 0",
                                 font=('Segoe UI', 10),
                                 bg='white')
        progress_label.pack(pady=5)
        
        # Log area
        log_frame = tk.Frame(progress_win, bg='white')
        log_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        log_text = scrolledtext.ScrolledText(log_frame,
                                            font=('Consolas', 9),
                                            bg='#f8fafc',
                                            relief='flat')
        log_text.pack(fill='both', expand=True)
        
        # Buttons
        btn_frame = tk.Frame(progress_win, bg='white')
        btn_frame.pack(fill='x', padx=20, pady=15)
        
        cancel_btn = tk.Button(btn_frame,
                              text="⏹ Cancel",
                              command=lambda: self.cancel_sending(progress_win),
                              font=('Segoe UI', 10, 'bold'),
                              bg=self.colors['danger'],
                              fg='white',
                              cursor='hand2',
                              relief='flat',
                              padx=20,
                              pady=8)
        cancel_btn.pack()
        
        # Start sending in thread
        threading.Thread(
            target=self.send_emails_thread,
            args=(email_list, log_text, progress_label, cancel_btn, progress_win, mode),
            daemon=True
        ).start()
    
    def send_emails_thread(self, email_list, log_text, progress_label, cancel_btn, progress_win, mode):
        """Thread function to send emails"""
        sent_count = 0
        failed_count = 0
        
        try:
            # Get settings
            sender_email = self.gmail_email.get().strip()
            sender_password = self.gmail_password.get().strip()
            
            # Get subject and body based on mode
            if mode == 'bulk':
                subject = self.bulk_subject.get().strip()
                body_html = self.get_html_body(self.bulk_body)
            else:  # individual
                subject = self.individual_subject.get().strip()
                body_html = self.get_html_body(self.individual_body)
            
            # Connect to SMTP
            log_text.insert('end', "🔌 Connecting to Gmail SMTP server...\n")
            log_text.see('end')
            
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)
            
            log_text.insert('end', "✅ Connected successfully!\n\n")
            log_text.see('end')
            
            # Send to each recipient
            for idx, recipient in enumerate(email_list, 1):
                if self.stop_sending:
                    log_text.insert('end', "\n⚠️ Sending cancelled by user.\n")
                    break
                
                try:
                    # Create message
                    msg = EmailMessage()
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = subject
                    msg.add_alternative(body_html, subtype='html')
                    
                    # Add attachments
                    for attachment_path in self.attachments:
                        with open(attachment_path, 'rb') as f:
                            file_data = f.read()
                            file_name = os.path.basename(attachment_path)
                            
                            # Guess MIME type
                            mime_type, _ = mimetypes.guess_type(attachment_path)
                            if mime_type:
                                maintype, subtype = mime_type.split('/', 1)
                            else:
                                maintype, subtype = 'application', 'octet-stream'
                            
                            msg.add_attachment(file_data,
                                             maintype=maintype,
                                             subtype=subtype,
                                             filename=file_name)
                    
                    # Send
                    server.send_message(msg)
                    sent_count += 1
                    log_text.insert('end', f"✅ [{idx}/{len(email_list)}] Sent to: {recipient}\n")
                    
                except Exception as e:
                    failed_count += 1
                    log_text.insert('end', f"❌ [{idx}/{len(email_list)}] Failed to: {recipient}\n    Error: {str(e)}\n")
                
                # Update progress
                progress_label.config(text=f"Total: {len(email_list)} | Sent: {sent_count} | Failed: {failed_count}")
                log_text.see('end')
            
            server.quit()
            
            if not self.stop_sending:
                log_text.insert('end', f"\n{'='*50}\n")
                log_text.insert('end', f"✅ Completed! Sent: {sent_count}, Failed: {failed_count}\n")
            
        except Exception as e:
            log_text.insert('end', f"\n❌ Error: {str(e)}\n")
            messagebox.showerror("Error", f"Failed to send emails:\n{str(e)}")
        
        finally:
            cancel_btn.config(text="Close", command=progress_win.destroy)
            log_text.see('end')
    
    def cancel_sending(self, progress_win):
        """Cancel ongoing email sending"""
        if self.stop_sending:
            progress_win.destroy()
        else:
            if messagebox.askyesno("Confirm", "Stop sending emails?"):
                self.stop_sending = True


def main():
    root = tk.Tk()
    app = EmailSenderPro(root)
    root.mainloop()


if __name__ == '__main__':
    main()
