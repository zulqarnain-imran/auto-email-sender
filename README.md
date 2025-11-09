# Professional Email Sender 📧

A modern, professional desktop email client built with **Python + Tkinter** for sending bulk and individual emails with attachments. Features a beautiful tabbed interface with file upload, column selection, and rich-text formatting.

---

## ✨ Features

### 📤 Bulk Email Sending
* **📁 Upload email lists** from Excel (.xlsx) or CSV files
* **🎯 Dynamic column selection** - choose which column contains emails from any file structure
* **👁️ Preview email addresses** before sending with email count
* **📊 Flexible file format** - works with any .xlsx or .csv file structure
* **🔄 Auto-detects columns** with "email" in their name

### ✉️ Individual Email Sending
* **📧 Send to individual recipients** with manual email entry
* **✅ Built-in email validation** to ensure valid addresses
* **💬 Perfect for quick one-off emails** without needing a file

### 📎 Advanced Attachment Support
* **📄 Attach multiple files** of any type (PDF, DOCX, TXT, images, etc.)
* **🗂️ Visual file management** with scrollable list and remove buttons
* **➕ Unlimited attachments** - add as many files as you need
* **🔄 Smart MIME type detection** for all file types
* **📎 Shared attachments** - same attachments apply to all modes

### 🎨 Professional User Interface
* **🌟 Modern tabbed design** - Bulk Email | Individual Email | Gmail Settings
* **💅 Beautiful card-based layout** with professional color scheme
* **📱 Responsive design** that adapts to window resizing
* **🎯 Clear visual feedback** for all actions
* **⚡ Intuitive navigation** with emoji icons
* **🖥️ Full-screen mode** with maximized window on startup

### ✨ Rich Text Formatting
* **📝 Bold and italic formatting** (Ctrl+B and Ctrl+I shortcuts)
* **🖊️ Visual formatting toolbar** for easy text styling
* **📧 HTML email support** - formatted text displays correctly in recipient's inbox
* **🎨 WYSIWYG editor** - see formatting as you type
* **🔤 Separate formatting** for bulk and individual email modes

### 📊 Real-time Progress Tracking
* **✅ Live status window** with real-time updates
* **📈 Progress counters** - Total, Sent, Failed counts
* **📋 Detailed logs** for each email sent with timestamps
* **⏹️ Cancel anytime** - stop sending mid-process
* **🔍 Error details** for troubleshooting failed emails
* **📊 Visual progress** with console-style logging

### 🔐 Secure & Reliable
* **🔑 Gmail App Password** support for secure authentication
* **🛡️ No password storage** - credentials only in memory during session
* **🔒 TLS encryption** via Gmail SMTP (port 587)
* **✅ Connection validation** before sending
* **🔒 Secure SMTP** with STARTTLS protocol

### 🚀 Easy Setup & Use
* **📦 One-click startup** via `start.py` (auto-installs all dependencies)
* **🖥️ Cross-platform** - Windows, Linux, and macOS
* **⚡ Zero configuration** needed
* **🎓 User-friendly** - no technical knowledge required
* **🔧 Automatic dependency management** - installs pandas and openpyxl automatically

---

## 📋 Requirements

* **Python 3.8+** (Download from [python.org](https://www.python.org/downloads/))
* **Gmail Account** with App Password enabled
* No manual package installation needed! `start.py` will automatically install:
  * `pandas==2.3.3` - for Excel/CSV file handling
  * `openpyxl>=3.0.0` - for .xlsx file support

All other required libraries (`tkinter`, `smtplib`, `email`, `mimetypes`, `threading`) are included with Python standard library.

---

## 🚀 Quick Start

### 1. Clone or Download the Repository

```bash
git clone https://github.com/zulqarnain-imran/auto-email-sender.git
cd auto-email-sender
```

### 2. Enable Gmail App Password

Before using this application, you need to create a Gmail App Password:

1. Go to [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification** if not already enabled
3. Search for "App passwords" in the search bar
4. Click **Create** and select **Mail** as the app type
5. You'll get a 16-character password like: `abcd efgh ijkl mnop`
6. Copy this password (remove spaces if needed)

⚠️ **Important:** Use this App Password in the application, NOT your regular Gmail password!

### 3. Run the Application

Simply double-click `start.py` or run:

```bash
python start.py
```

The script will:
- ✅ Check for required packages (pandas, openpyxl)
- 📦 Auto-install any missing dependencies
- 🚀 Launch the Professional Email Sender
- 🖥️ Open in full-screen mode

---

## 📖 How to Use

### Configure Gmail Settings (Do This First):

1. **Navigate to "🔐 Gmail Settings" tab**
2. **Enter your Gmail address** (e.g., yourname@gmail.com)
3. **Enter your Gmail App Password** (the 16-character password from Google)
4. These credentials will be used for all email sending

### For Bulk Email Sending:

1. **Navigate to "📤 Bulk Email" tab**
2. **Click "📂 Choose File"** and select your .xlsx or .csv file
3. **Select the email column** from the dropdown menu (columns with "email" are auto-suggested)
4. **Preview the email addresses** to verify they're correct (shows up to 50 addresses)
5. **Compose your email:**
   - Enter subject in the Subject field
   - Write your message in the Message Body field
   - Use formatting toolbar (B for bold, I for italic)
   - Press Ctrl+B for bold, Ctrl+I for italic
6. **Add attachments (optional):**
   - Click "📎 Add" to browse and select files
   - Multiple files can be added
   - Remove individual files with ✖ button or clear all
7. **Click "🚀 Send Bulk Emails"**
8. **Monitor progress** in the popup window with live updates

### For Individual Email:

1. **Navigate to "✉️ Individual Email" tab**
2. **Enter the recipient's email address** (validated automatically)
3. **Compose your email:**
   - Enter subject in the Subject field
   - Write your message in the Message Body field
   - Use formatting toolbar for bold/italic text
4. **Add attachments if needed** (same as bulk mode)
5. **Click "✉️ Send Email"**
6. **Monitor progress** in the popup window

### Formatting Your Email Body:

- **Bold text:** Select text and press `Ctrl+B` or click the "B" button
- **Italic text:** Select text and press `Ctrl+I` or click the "I" button
- Text formatting will be preserved as HTML in the sent email
- Formatting works independently in both Bulk and Individual modes

---

## 📁 File Format

Your email list file (.xlsx or .csv) can have **any structure**. For example:

### Example 1: Simple email list
| Email               |
|---------------------|
| hr1@company.com     |
| hr2@company.com     |
| hr3@company.com     |

### Example 2: Multiple columns
| Name    | Email               | Company  | Position |
|---------|---------------------|----------|----------|
| John    | john@company.com    | ABC Corp | HR       |
| Sarah   | sarah@company.com   | XYZ Inc  | Manager  |

The app will let you **select which column** contains the email addresses!

---

## 🎯 Key Features Explained

### Dynamic Column Selection
The app provides flexibility with email list files:
- Upload any Excel (.xlsx) or CSV file with any structure
- Dropdown menu shows all available columns from your file
- Select which column contains the email addresses
- Auto-detection: columns with "email" in the name are automatically suggested
- Preview shows up to 50 email addresses before sending

### Multi-File Attachments
Powerful attachment management:
- **Any file type** - PDF, DOCX, TXT, PNG, JPG, ZIP, etc.
- **Multiple files** - attach unlimited files (within email size limits)
- **Easy management** - scrollable visual list with individual remove buttons
- **Shared across modes** - attachments added in one mode are available in both
- **Smart MIME detection** - automatically detects file types
- **Clear all option** - remove all attachments with one click

### Modern UI Design
Built with user experience in mind:
- **Tabbed interface** - organized into 3 logical sections (Bulk, Individual, Settings)
- **Card-based layout** - clean, modern appearance with visual hierarchy
- **Professional color scheme** - primary blue (#2563eb), secondary green (#10b981), danger red (#ef4444)
- **Full-screen** - launches in maximized mode for better workspace
- **Responsive** - adapts to window resizing
- **Emoji icons** - visual indicators for better UX and quick navigation
- **Scrollable areas** - for attachment lists and email previews

### Threading & Progress Tracking
Advanced progress monitoring:
- **Non-blocking UI** - sending happens in background thread
- **Real-time updates** - live status of Total/Sent/Failed counts
- **Detailed logging** - console-style log with success/failure for each email
- **Cancel anytime** - stop button to abort sending mid-process
- **Error reporting** - detailed error messages for debugging

---

## 🛠️ Troubleshooting

### "Authentication failed" or Login errors
- Make sure you're using a Gmail **App Password**, not your regular password
- Verify 2-Step Verification is enabled on your Google account
- Check that the email address is correct and matches the App Password account
- Remove any spaces from the App Password
- Try generating a new App Password

### "No file selected" or File upload errors
- Click "📂 Choose File" and select a valid .xlsx or .csv file
- Make sure the file isn't corrupted or password-protected
- Verify the file has at least one column with email addresses
- Try opening the file in Excel/Sheets first to verify it's readable

### "Missing Info" warnings
- Navigate to "🔐 Gmail Settings" tab and fill in all required fields
- Gmail address and App Password are required
- For sending emails, subject and body are also required
- The app will automatically guide you to the correct tab

### Attachments not sending or disappearing
- Verify the attachment file still exists at the original location
- Don't move or delete files after adding them as attachments
- Check file permissions (ensure file isn't locked by another program)
- Try re-adding the attachment if it was removed
- Large attachments may cause timeouts - keep total size under 25MB

### "Failed to load file" or Column selection issues
- Ensure the file is a valid .xlsx or .csv format
- Try opening the file in Excel/Sheets to verify it's not corrupted
- Check that pandas and openpyxl are installed (run `python start.py`)
- Make sure the file has proper headers in the first row
- Verify the selected column actually contains email addresses

### Application won't start
- Ensure Python 3.8+ is installed: `python --version`
- Run `python start.py` instead of `python email_sender_pro.py`
- Check if tkinter is installed: `python -c "import tkinter"`
- On Linux, install tkinter: `sudo apt-get install python3-tk`
- Check console for error messages

### Emails going to spam
- Ask recipients to check their spam/junk folder
- Add a clear subject line and professional email body
- Avoid spam trigger words in subject and body
- Send in smaller batches (50-100 at a time)
- Wait between batches to avoid Gmail rate limits

---

## 📁 Project Structure

```
auto-email-sender/
│
├── email_sender_pro.py    # Main application with GUI
├── start.py                # Startup script (auto-installs dependencies)
├── requirements.txt        # Python package dependencies
└── README.md              # This documentation file
```

### File Descriptions

- **`email_sender_pro.py`**: Core application with complete GUI implementation using Tkinter
- **`start.py`**: Launcher script that checks and installs required packages automatically
- **`requirements.txt`**: Lists all Python package dependencies (pandas, openpyxl)

---

## 🤝 Contributing

Contributions are welcome! Feel free to:
- 🐛 Report bugs via [Issues](https://github.com/zulqarnain-imran/auto-email-sender/issues)
- 💡 Suggest new features or improvements
- 🔧 Submit pull requests with enhancements
- 📖 Improve documentation
- ⭐ Star the repository if you find it useful

### Development Setup
1. Fork the repository
2. Clone your fork: `git clone https://github.com/YOUR_USERNAME/auto-email-sender.git`
3. Create a feature branch: `git checkout -b feature-name`
4. Make your changes and test thoroughly
5. Commit: `git commit -m "Add feature description"`
6. Push: `git push origin feature-name`
7. Create a Pull Request

---

## 📄 License

This project is free to use for personal and educational purposes. Feel free to modify and distribute as needed.

---

## 👨‍💻 Author

**Zulqarnain Imran**
- GitHub: [@zulqarnain-imran](https://github.com/zulqarnain-imran)
- Repository: [auto-email-sender](https://github.com/zulqarnain-imran/auto-email-sender)

---

## 💡 Tips & Best Practices

1. **Test with small lists first** - send to 2-3 test addresses to verify everything works correctly
2. **Keep attachments reasonable** - total size should be under 25MB to avoid timeouts
3. **Use descriptive subjects** - clear subjects help recipients identify your email
4. **Preview emails** - always check the preview before bulk sending (shows up to 50 addresses)
5. **Gmail sending limits** - Free Gmail accounts can send up to 500 emails/day
6. **Monitor the progress window** - watch for any failed sends and check error messages
7. **Format text carefully** - use bold/italic sparingly for professional appearance
8. **Verify email addresses** - preview ensures all addresses are valid before sending
9. **Don't close progress window** - wait for completion to see full results
10. **Add attachments carefully** - don't move files after adding them to the list
11. **Batch large lists** - for 500+ emails, split into multiple batches over multiple days
12. **Professional email structure** - use clear greeting, body, and signature for best results

---

## 📞 Support

If you encounter issues:
1. Check the **Troubleshooting** section above for common problems
2. Verify your Gmail App Password is correct and 2FA is enabled
3. Ensure Python 3.8+ is installed: `python --version`
4. Run `python start.py` to auto-fix dependency issues
5. Open an [Issue](https://github.com/zulqarnain-imran/auto-email-sender/issues) on GitHub with:
   - Detailed description of the problem
   - Error messages or screenshots
   - Your Python version and OS

---

## 🌟 Show Your Support

If you find this project helpful:
- ⭐ Star the repository on GitHub
- 🍴 Fork it to customize for your needs
- 📢 Share it with others who might benefit
- 🐛 Report bugs to help improve it

---

**Made with ❤️ using Python + Tkinter**

*Perfect for job applications, bulk email campaigns, newsletters, and professional communications!*
