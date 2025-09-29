# Clik

A smart Outlook add-in that converts emails to calendar events using AI-powered analysis and generates direct Outlook calendar links.

## 🔧 Setup

1. Clone the repo:
   ```bash
   git clone https://github.com/blip-cmd/Clik.git
   ```

2. Install dev certificate with [mkcert](https://github.com/FiloSottile/mkcert):
   ```bash
   mkcert localhost
   ```

3. Serve task pane over HTTPS:
   ```bash
   node ClikWeb/server.js
   ```

4. Sideload the add-in in Outlook:
   - Open Outlook Web or Desktop
   - Go to **Settings > Manage Add-ins** ([link](https://aka.ms/olksideload))
   - Upload the manifest file from `/ClikManifest/`

## ✅ Features

- **AI-Powered Email Analysis**: Uses Google Gemini API to extract event details from email content
- **Smart Calendar Links**: Generates direct Outlook calendar links with pre-filled event data
- **Multiple Export Options**: 
  - Generate calendar links for Outlook Web
  - Download .ics files for any calendar app
- **Clean Task Pane Interface**: User-friendly interface with copy/paste functionality
- **Console Link Display**: Always shows full calendar link in browser console
- **Secure HTTPS Serving**: CSP-compliant asset loading

## 📁 Structure

```
Clik/
├── .git/                   # Git version control
├── .vs/                    # Visual Studio settings
├── Clik/                   # Core add-in logic and manifest
├── ClikWeb/                # Web project for task pane and functions
├── packages/               # NuGet packages
├── localhost.pem           # HTTPS certificate (mkcert)
├── localhost-key.pem       # HTTPS key (mkcert)
├── Clik.sln                # Visual Studio solution file
├── .gitignore              # Git ignore rules
└── readme.md               # Project documentation
```

## 🚀 How to Use

1. **Select an email** in Outlook that contains event information
2. **Click "Calendar Link"** button in the Clik ribbon group
3. **Task pane opens** with generated calendar link
4. **Copy the link** or click "Open in Outlook" to create the event
5. **Alternative**: Use "Download .ics" for traditional calendar file download

```
## 📌 Notes

- Tested on Outlook Web
- Requires HTTPS for CSP compliance
- Ideal for demo, prototyping, or AppSource prep

---

Built with ❤️ by Ryan — optimizing workflows, one click at a time.
