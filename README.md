# 🎓Automated Certificate Generation and Emailing System using UiPath

---

## 🚀 Overview

This UiPath automation project simplifies the manual task of sending participation certificates to webinar/workshop attendees. The bot reads participant details from an Excel sheet, generates personalized PDF certificates from a PowerPoint template, sends them via email using SMTP, and logs the status back into the same Excel file.

✅ Built for event coordinators, educators, and admins who frequently manage certificate distribution.

---

## 🧠 Problem Statement

Avirup, an event coordinator, needs to send certificates after each event. Doing this manually is time-consuming ⏳ and error-prone ❌. 

💡 This bot solves that by:
- 📥 Reading names and emails from Excel
- 📄 Generating personalized certificates in PDF
- 📧 Emailing the certificate
- 📊 Updating the Excel with status ("✅ Sent"/"❌ Failed") and timestamp

---

## ✨ Features

- 📊 Reads participant data (Name, Email) from Excel
- 🧾 Replaces name field in PowerPoint template
- 📤 Saves certificate as PDF
- 📧 Sends email with the certificate attached
- 🔁 Loops through all entries
- 🕓 Writes "Sent"/"Failed" and timestamp back to Excel
- 🛡️ Uses Try-Catch for error handling

---

## 🛠️ Technologies Used

- **UiPath Studio** (Tested on 2023.10+)
- Excel & PowerPoint Integration
- SMTP Email Activity
- VB.NET (for text replacement and data conversion)
- Try-Catch Exception Handling

---

---

## 📝 How to Use

1. Clone or download this repository.
2. Open `Main.xaml` in UiPath Studio.
3. Modify paths inside:
   - Excel Application Scope → Excel file
   - PowerPoint Template → Your certificate template
   - Save PDF location → Output folder
4. Configure email settings in the **Send SMTP Mail** activity:
   - Use Gmail/Outlook SMTP and app password if needed.
5. Run the bot.
6. Check Excel for status updates and timestamps.

---

## 🧪 Test Data Format (ParticipantList.xlsx)

| NAME        | EMAIL                | STATUS | TIMESTAMP          |
|-------------|----------------------|--------|---------------------|
| John Doe    | john@example.com     | Sent   | 2025-08-02 11:15:00 |
| Priya Shah  | priya@example.com    | Failed | 2025-08-02 11:17:20 |

---

## 🧩 Customization Ideas

- Add event name/date on the certificate
- Use HTML emails for better formatting
- Integrate with Orchestrator for scheduling
- Store certificate links in Excel instead of attachments

---

## 📌 Best Practices

- 🔐 Don’t hardcode credentials – use UiPath Assets or Windows Credential Store
- 🔁 Use Try-Catch to handle email failures
- 🔍 Add logging via `Log Message` or `Write Line`

---

## 🏁 Outcome

- ⏱️ Saves time by eliminating repetitive tasks
- 📈 Reduces human errors
- ✅ Delivers certificates with 100% consistency

---

## 🤝 Contributing

Pull requests are welcome! For major updates, open an issue first to discuss changes.

---

## 📃 License

This project is licensed under the MIT License.

---

## 📂 Folder Structure
