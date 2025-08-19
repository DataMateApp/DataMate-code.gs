# DataMate Code for Google Sheets  
**Transform your spreadsheets into powerful data management tools.**  

## About DataMateApps  
Hi, I'm **Dan Northway**—Founder and Developer of DataMateApps. Before retiring, I spent my career as a **Construction Project Manager and Superintendent**.  

In construction, spreadsheets were the backbone of project management—timecards, pay applications, logs—all stored in countless Excel files and PDFs, tracked manually. I knew databases could streamline this, but Excel was the industry standard. That sparked an idea:  
> *What if I could turn Excel into a lightweight database?*  

Using **forms and VBA**, I built a system to store, log, and organize data efficiently—making sorting and filtering effortless. It became indispensable in my workflow.  

### The Birth of DataMate  
Post-retirement, I revisited this concept with a new twist:  
> *Why not make it work with any form?*  

With help from **ChatGPT**, **DataMate was born!** Built for **Google Sheets** using Apps Script, it’s tailored for **small businesses and teams** who rely on spreadsheets but need a **smarter, structured way to manage data**. It’s not a replacement for **SQL databases** or enterprise systems—it bridges the gap between **manual spreadsheets** and **complex (often costly) solutions** that may overwhelm smaller operations.  

### Why is DataMate Free?  
Because **the idea matters more than the programming.**  

Technology has democratized development, and for me, this is a **passion project** to **keep my mind sharp** and leave a **legacy**. DataMate grows with every user—your feedback shapes its future!  

---

### Installation and Deployment
[**DataMate Open Source Template**](https://docs.google.com/spreadsheets/d/1G-zoZx6OT4DhdA-yAPI--ZGLDPynkaFLdfRjU4RAX-Q/template/preview) 

Open-source code to deploy as web app.

Or:
#### Step 1: Install the Core Script  
1. Open **Google Drive** and create a **new Google Sheet**.  
2. Click **Extensions > Apps Script**.  
3. Delete any default code in `Code.gs`.  
4. Copy & paste the [`Code.gs`](https://github.com/DataMateApp/DataMate-code.gs) from this repository.  
5. Click **Save** (Ctrl+S or the disk icon).  

#### Step 2: Add Supporting HTML Files  
DataMate uses HTML files to enhance functionality: `tutorial.html`, `FormBuilder.html`, and (optionally) `UploadCSV.html`.  
1. In the Apps Script editor, click the **+** button next to "Files" and select **HTML**.  
2. Create `tutorial.html`:  
   - Name it `tutorial.html`.  
   - Copy the contents from [this repository’s `help.html`](https://datamateapp.github.io/help.html) or the updated version in the DataMate documentation.  
   - This displays a modal tutorial (via `showTutorial()`).  
3. Create `FormBuilder.html`:  
   - Name it `FormBuilder.html`.  
   - Copy its contents from the repository (or implement a drag-and-drop UI if available).  
   - This powers the visual form editor (via `showFormBuilder()`).
4. Create `MailIt.html`:  
   - Name it `MailIt.html`.  
   - Copy its contents from the repository (or implement a drag-and-drop UI if available).  
   - This powers the visual form editor (via `showMailItSidebar()`).   
5. (Optional) Create `UploadCSV.html`:  
   - Name it `UploadCSV.html`.  
   - Copy its contents from the repository.  
   - This enables batch CSV uploads for data entry.  
6. Click **Save** for each file.  

> **Note**: If `FormBuilder.html` or `UploadCSV.html` aren’t in the repository, basic placeholders can be created (e.g., a form for CSV upload or a field editor). Contact [datamateapp@gmail.com](mailto:datamateapp@gmail.com) for assistance.

#### Step 3: Initialize and Test  
1. In the Apps Script editor, click **Run > showTutorial** to test permissions.  
2. Authorize the script (grant access to Google Sheets and Drive when prompted).  
3. Open your spreadsheet, refresh it (F5), and look for the **DataMate** menu.  
4. Select **DataMate > FormBuilder > Preview Form** to initialize the `FormSetup` sheet with sample fields (e.g., Text, Checkout, Hyperlink).  
   - This creates a pre-configured `FormSetup` sheet starting at `A9:J` with 29 field examples.  

#### Step 4: Web Deployment (Optional)  
Deploy DataMate as a web app to share forms with others:  
1. In the Apps Script editor, click **Deploy > New Deployment**.  
2. Select **Web App**.  
3. Configure:  
   - **Description**: E.g., "DataMate FormBuilder".  
   - **Execute as**: "Me" (runs under your account).  
   - **Who has access**: "Anyone" (public) or "Anyone with a Google account" (Google users only).  
4. Click **Deploy** and copy the **Web App URL**.  
5. Share the URL for users to access forms directly in their browsers.  
   - Example: Deploy `generateFormHTML()` (via `doGet(e)`) to serve the form defined in `FormSetup`.  
6. To update, go to **Deploy > Manage Deployments**, select your deployment, and click **New Version**.  

> **Tip**: The provided `Code.gs` includes `doGet(e)` to serve `generateFormHTML()`. Test the URL in a browser to preview your form.

---

### Completed Features  
- **Spreadsheet Database**: Store and organize data across sheets (e.g., `Responses`, `Input`).  
- **FormBuilder**: Build custom forms with **29 field types**, including:  
  - **Basic**: Text, Dropdown, Checkbox, Radio, Textarea, Email, Number, Date, Time  
  - **Advanced**: StarRating, RangeSlider, FileUpload (6MB max), Signature, Geolocation  
  - **Dynamic**: Conditional (show/hide logic), Calculated (basic formulas like `=Number*2`), Checkout (order tables with quantities and totals), Hyperlink (clickable links)  
  - **Display**: StaticText, Table (renders sheet ranges with images/videos), Header/Footer (HTML support)  
  - **Media**: Image, Video, ImageLink, VideoLink  
  - **Utility**: ProgressBar, Captcha (fixed "3 + 5 = 8")  
  - **Layout**: Container (styled grouping)  
- **Data Logging**: Map form inputs to sheets via `FormSetup!B:G` (up to 3 targets per field).  
- **Email Notifications**: Send automated emails with form response details on submission (configured via `FormSetup!B8`).  
- **UI**: Simple form preview (`previewForm()`) and visual editor (`showFormBuilder()`).  
- **File Uploads**: Stores files in Google Drive with public links.  
- **Custom Actions**: Run functions post-submission (e.g., `save, newcontact`) from `FormSetup!B6`.  
- **Tutorial**: In-app guide via `tutorial.html`.  

### Upcoming Features  
- **Multi-User Collaboration**: Secure data access for teams.  
- **Automated Backups**: Snapshot data for recovery.  
- **Form Integrations**: Link with Google Forms or external sources.  
- **Advanced Filtering**: Enhanced search and sort tools.  
- **Collaboration Welcome!**: Need help to build these—email [datamateapp@gmail.com](mailto:datamateapp@gmail.com).  

**Have ideas?** Share them at [[datamateapp@gmail.com](mailto:datamateapp@gmail.com](https://script.google.com/macros/s/AKfycbyPu31wwkbtKZ5X-382wdKzG2Y8vIN-fSApo_lAR9x_1n_qgxPUgmAafmljS6RCc3i7/exec)) or visit [our website](https://datamateapp.github.io/).  

---

## 💙 Support This Project  

DataMate is free, but your support fuels its growth:  

[**Support DataMateApps**](https://datamateapp.github.io/Donate%205%20per%20mo.html)  

Every donation keeps this project thriving!  

---

## License  
Licensed under the **MIT License**. See [`LICENSE.txt`](https://github.com/DataMateApp/DataMate-code.gs/blob/main/LICENSE.txt) in the repository for details.  

## Credits  
Developed by **Dan Northway** Founder and **Sara Bohannon** Co-Founder. Special thanks to the open-source community and AI tools like ChatGPT for accelerating development.
