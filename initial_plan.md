Absolutely, Nick — here’s a **complete, ready‑to‑deliver 3‑day course** on becoming a SharePoint Developer, where students finish with a working **Azure-hosted web app that reads/writes documents to a SharePoint document library**.

I’ve structured it to balance **hands‑on labs**, **architecture fundamentals**, and **real‑world developer scenarios**.

***

# 🚀 3‑Day Course: *Becoming a SharePoint Developer*

### **Final Project:** Build an Azure Web App (Python/Flask or Node.js) that authenticates with Microsoft Entra ID and **reads/writes documents to SharePoint** via Microsoft Graph.

***

# **🔹 Day 1 — Foundations of SharePoint Development**

### **Theme:** Understand SharePoint as a platform + authentication basics + environment setup

***

## **📘 Module 1 — SharePoint Developer Landscape**

**Topics**

*   SharePoint Online architecture
*   Site collections, lists, libraries
*   SharePoint REST API vs **Microsoft Graph API**
*   What developers actually build:
    *   SPFx customisations
    *   Power Automate extensions
    *   External apps integrating via Graph

**Outcome:** Students understand where custom code interacts with SharePoint.

***

## **🧰 Module 2 — Developer Environment Setup**

**Stack options provided:**

*   Node.js + Express
*   Python + Flask
*   (Instructor may standardise—Flask is simple and great for teaching)

**Hands-on**

*   Create Microsoft 365 Developer Account (E5 dev tenant)
*   Register Azure AD (Entra ID) App
*   Configure API permissions:
    *   `Files.ReadWrite.All`
    *   `Sites.ReadWrite.All`
*   Generate client secret
*   Understand:
    *   OAuth 2.0 Authorization Code Flow
    *   Client secrets vs certificates

**Outcome:** Students have a working dev environment & registered app.

***

## **💻 Module 3 — Microsoft Graph Basics**

**Topics**

*   Making authenticated Graph calls
*   Access tokens, refresh tokens
*   Using Graph Explorer
*   Finding SharePoint site & library IDs

**Hands-on**

*   Use Graph Explorer to:
    *   list sites
    *   list drives
    *   list items in a library
    *   upload / download single file

**Outcome:** Students can query SharePoint from Graph manually.

***

## **🧪 Lab 1 — Build Starter Web App**

**Students create:**

*   A minimal Flask/Express app
*   Login with Azure AD (using MSAL)
*   Display user's profile information from Graph

**Outcome:** Basic working web app with authentication.

***

<br>

# **🔹 Day 2 — Reading & Writing SharePoint Files**

### **Theme:** Graph deep‑dive + application logic + storage operations

***

## **📘 Module 4 — Working with SharePoint Document Libraries**

**Topics**

*   Libraries vs Drives in Graph
*   DriveItems (files + folders)
*   Permissions models
*   Uploading:
    *   simple upload (<4MB)
    *   chunked/resumable upload (large files)

**Hands-on**

*   Find library drive ID using Graph
*   Use Graph SDK (or raw REST calls) to:
    *   Read metadata
    *   Upload file
    *   Download file
    *   Delete file
    *   Create folder

***

## **🧪 Lab 2 — Integrate SharePoint Storage in Web App**

Students add features to web app:

### **✔ View SharePoint Library Files**

*   list items using:

```http
GET /sites/{siteId}/drives/{driveId}/root/children
```

### **✔ Download a File**

Backend:

```python
GET /sites/{siteId}/drives/{driveId}/items/{itemId}/content
```

### **✔ Upload a File**

Frontend: simple HTML file upload  
Backend: POST to Graph upload endpoint

**Outcome:** Web app now reads & writes SharePoint docs successfully.

***

## **📦 Module 5 — App Registration Deep Dive**

*   App roles & delegated permissions
*   Admin consent
*   Authentication best practices
*   Storing credentials securely
*   Multi-tenant vs single-tenant apps

***

## **🧪 Lab 3 — Store App Secrets Securely**

Students:

*   Move secrets into **Azure Key Vault**
*   Update app to fetch secrets at runtime

**Outcome:** App is cloud-ready and secure.

***

<br>

# **🔹 Day 3 — Deploying & Productionising**

### **Theme:** Azure deployment + enterprise‑ready features + final project

***

## **📘 Module 6 — Hosting the Application on Azure**

**Topics**

*   Azure Web App basics
*   Deployment options:
    *   VS Code GitHub Actions
    *   Azure DevOps
    *   Zip deploy / container deploy
*   Setting environment variables in Azure
*   Aligning redirect URIs with Azure AD app

**Outcome:** Students understand deployment pipeline.

***

## **🧪 Lab 4 — Deploy Web App to Azure**

Steps:

1.  Create Azure Web App
2.  Configure runtime (Node/Python)
3.  Set environment variables:
    *   CLIENT\_ID
    *   CLIENT\_SECRET
    *   TENANT\_ID
    *   KEYVAULT\_URL
4.  Deploy via GitHub Actions or VS Code
5.  Test authentication
6.  Validate SharePoint file operations in production

**Outcome:** The app is live on the internet.

***

## **📘 Module 7 — Polishing & Extending the App**

Options for enhancement:

*   Preview documents with Graph DriveItem thumbnails
*   Add folder creation
*   Add version history view
*   Logging via Application Insights
*   Pagination for large libraries
*   Error handling patterns:
    *   token expiry
    *   permission denials
    *   throttling (HTTP 429)

***

## **🧪 Final Project — Build a Production‑Ready SharePoint Document Manager**

Students build a fully functioning **SharePoint Document Manager Web App** hosted on Azure:

### **Core Features**

✔ Secure login using Microsoft identity  
✔ Browse SharePoint document library  
✔ Upload files  
✔ Download files  
✔ Delete files  
✔ Create folders  
✔ Log activity

### **Deliverables**

*   Source code in GitHub
*   Running Azure Web App URL
*   README documenting architecture

***

# 🎓 **End-of-Course Outcomes**

Students will be able to:

### 🔧 Technical Skills

*   Build modern web apps integrating with Microsoft 365
*   Use MSAL for authentication
*   Use Microsoft Graph API for SharePoint
*   Manage permissions & Azure AD apps
*   Deploy secure apps on Azure Web Apps
*   Store and use secrets via Azure Key Vault

### 📚 SharePoint Development Skills

*   Understand site structures & document libraries
*   Perform CRUD operations on files via Graph
*   Navigate SharePoint’s REST & Graph models

### ☁ Cloud & DevOps Skills

*   Deploy cloud apps to Azure
*   Configure CI/CD via GitHub or Azure DevOps

***

If you'd like, I can also generate:

✅ slide deck for all 3 days  
✅ student handouts  
✅ code templates (Flask or Node)  
✅ full final project repo structure  
✅ certification‑style quiz questions

Would you like those next?
