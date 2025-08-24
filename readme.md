# OutlookWebAddIn1

A lightweight Outlook add-in that automates email-to-calendar workflows via a secure local dev server.

## 🔧 Setup

1. Clone the repo:
   ```bash
   git clone https://github.com/your-username/OutlookWebAddIn1.git
   ```

2. Install dev certificate with [mkcert](https://github.com/FiloSottile/mkcert):
   ```bash
   mkcert localhost
   ```

3. Serve task pane over HTTPS:
   ```bash
   npx http-server ./dist -S -C localhost.pem -K localhost-key.pem
   ```

4. Sideload the add-in in Outlook:
   - Open Outlook Web or Desktop
   - Go to **Settings > Manage Add-ins**
   - Upload the manifest file from `/manifest/`

## ✅ Features

- Secure task pane rendering via HTTPS
- ICS generation for calendar automation
- CSP-compliant asset loading

## 📁 Structure

```
OutlookWebAddIn1/
├── .git/                   # Git version control
├── .vs/                    # Visual Studio settings
├── OutlookWebAddIn1/       # Core add-in logic and assets
├── OutlookWebAddIn1Web/    # Web project for task pane
├── packages/               # NuGet packages
├── localhost.pem           # HTTPS certificate (mkcert)
├── localhost-key.pem       # HTTPS key (mkcert)
├── OutlookWebAddIn1.sln    # Visual Studio solution file
├── .gitignore              # Git ignore rules
└── readme.md               # Project documentation
```

```

## 📌 Notes

- Tested on Outlook Web
- Requires HTTPS for CSP compliance
- Ideal for demo, prototyping, or AppSource prep

---

Built with ❤️ by Ryan — optimizing workflows, one click at a time.

