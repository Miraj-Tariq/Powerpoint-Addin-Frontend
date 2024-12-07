# 📝 PowerPoint Slide Manager Add-in

This project is a **PowerPoint Context Add-in** built with **React** and **Office.js** to manage and search slides. The add-in allows users to open a popup window, search for slides, and add them to a deck.

---

## 📋 Table of Contents

1. [Prerequisites](#prerequisites)
2. [Project Setup](#project-setup)
3. [Running the Project](#running-the-project)
4. [PowerPoint Configuration](#powerpoint-configuration)
5. [Testing the Add-in](#testing-the-add-in)
6. [Project Structure](#project-structure)
7. [Dependencies](#dependencies)
8. [Troubleshooting](#troubleshooting)

---

## 🚀 Prerequisites

Make sure you have the following tools installed:

1. **Node.js (v16.x or higher)**: [Download Node.js](https://nodejs.org/)
2. **npm (v7.x or higher)**: Comes with Node.js installation.
3. **PowerPoint (Office 365 or 2019)** with support for add-ins.

---

## 🛠 Project Setup

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/your-username/powerpoint-slide-manager.git
   cd powerpoint-slide-manager
   ```

2. **Install Dependencies**:
    ```bash
   npm install
   ```
   
3. **Install Type Definitions for Office.js (Optional for TypeScript)**:
    ```bash
   npm install --save-dev @types/office-js
   ```
    
## ▶️ Running the Project
1. **Enable HTTPS** for your local development server:

    Create a `.env` file in the root of the project with the following content:
    ```
    HTTPS=true
    ```    
2. **Start the React Development Server**:
    ```bash
    npm start
    ```
   This will serve the application at:
    ```
    https://localhost:3000
    ```
3. **Trust the Self-Signed Certificate**:

    When you start the server for the first time, you might get a security warning. Follow these steps:
   - Open https://localhost:3000 in your browser.
   - Accept the self-signed certificate.

## 🧩 PowerPoint Configuration
### Load the Add-in in PowerPoint
1. **Create a Network-Trusted Manifest**:

    Ensure your manifest.xml is correctly set up. Key sections:
    ```xml
    <DefaultSettings>
      <SourceLocation DefaultValue="https://localhost:3000/index.html"/>
    </DefaultSettings>
    
    <AppDomains>
      <AppDomain>https://localhost:3000</AppDomain>
    </AppDomains>
    ```
2. **Insert the Add-in in PowerPoint**:
   - Open PowerPoint.
   - Go to Insert > My Add-ins > Upload My Add-in.
   - Select the manifest.xml file.
3. **Verify the Add-in**:
   - Go to the Home tab.
   - Click the PowerPoint AI ChatBot button to open the task pane.
---
## Testing the Add-in
1. **Run the Add-in**:
   - Click "Search Slides" in the task pane.
   - Confirm that a popup appears asking for permission to open a new window.
   - Click "Allow".
2. **Verify the Popup**:
   - Ensure the popup displays the search bar, tabs, images grid (3 images per row), and pagination.
   - Test the "Add to Deck" button to confirm functionality.
---
## 📂 Project Structure
    ```
    powerpoint-slide-manager/
    ├── public/
    │   ├── dialog.html        # HTML file for the popup dialog
    │   └── index.html         # Main entry point for the React app
    ├── src/
    │   ├── components/
    │   │   ├── Header.tsx     # Header component
    │   │   ├── MainPage.tsx   # Main page with buttons and handlers
    │   │   └── Popup.tsx      # Popup component (not used externally)
    │   ├── App.tsx            # Main App component
    │   └── index.tsx          # ReactDOM render
    ├── manifest.xml            # PowerPoint add-in manifest
    └── package.json            # Project dependencies and scripts
    ```
---
## 📦 Dependencies
Key dependencies used in this project:
- **React**: JavaScript library for building user interfaces.
- **Material-UI**: UI component library for React.
- **Office.js**: Microsoft Office JavaScript API.

Install dependencies:

```bash
  npm install react @mui/material @mui/icons-material
```
Install type definitions (for TypeScript):

```bash
  npm install --save-dev @types/react @types/office-js
```