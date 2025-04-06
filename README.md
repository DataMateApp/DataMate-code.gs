# DataMate Code for Google Sheets  
**Transform your spreadsheets into powerful data management tools.**  

## About DataMateApps  
Hi, I'm **Dan Northway**—Founder and Developer of DataMateApps. Before retiring, I spent my career as a **Construction Project Manager and Superintendent**.  

At most companies I worked with, spreadsheets were the backbone of project management. Everything—timecards, pay applications, logs—was stored in countless Excel and PDF files, requiring manual tracking.  

I knew databases could streamline this process, but Excel was the standard. That’s when I had an idea:  
> *What if I could turn Excel itself into a lightweight database?*  

Using **forms and VBA**, I built a system that stored, logged, and organized data efficiently—making sorting and filtering a breeze. It became an essential tool in my workflow.  

### The Birth of DataMate  
After retiring, I revisited the concept, and a lightbulb went off:  
> *Why not make this work with any form?*  

With help from the recently released **ChatGPT**, **DataMate was born!**  

DataMate isn’t a replacement for **full-scale databases like SQL** or enterprise-level solutions. Instead, it’s designed for **small businesses and teams** that rely on spreadsheets but need a **smarter, structured way to manage data**.  

It bridges the gap between **manual spreadsheets** and **complex (often expensive) systems** that may be overkill for smaller operations.  

### Why are DataMateApps Free?  
Because **the idea matters more than the programming.**  

Technology has made development more accessible, and for me, this is both a **passion project** and a way to **keep my mind sharp**. More importantly, I see DataMate as a **legacy—one that grows and evolves with every user.**  

---

### Installation and Deployment  

#### Step 1: Install the Core Script  
1. Open **Google Drive** and create a **new spreadsheet**.  
2. Click **Extensions > Apps Script**.  
3. Delete any default code in `Code.gs`.  
4. Copy & paste the provided `Code.gs` from this repository.  
5. Click **Save**.  

#### Step 2: Add Supporting HTML Files  
DataMate includes optional HTML files to enhance functionality: `tutorial.html` and `UploadCSV.html`.  
1. In the Apps Script editor, click the **+** button next to "Files" and select **HTML**.  
2. Name the first file `tutorial.html`.  
   - Copy the contents of `tutorial.html` from this repository and paste it into the editor.  
   - This file provides a detailed guide on using DataMate, including FormBuilder features.  
3. Name the second file `UploadCSV.html`.  
   - Copy the contents of `UploadCSV.html` from this repository and paste it into the editor.  
   - This file enables batch CSV uploads to streamline data entry.  
4. Click **Save** for each file.  

#### Step 3: Deploy the Add-on  
1. In the Apps Script editor, click **Run > onInstall**.  
2. Authorize the script when prompted (you’ll need to grant permissions for Google Sheets and Drive access).  
3. Open your spreadsheet, refresh it, and look for the **DataMate** menu.  
4. Select **DataMate > New Dataset** to initialize your spreadsheet with specialized sheets (e.g., `Input`, `Data`, `View_Print`).  

#### Step 4: Web Deployment (Optional)  
DataMate can be deployed as a web app to share forms or tools with others:  
1. In the Apps Script editor, click **Deploy > New Deployment**.  
2. Choose **Web App** as the deployment type.  
3. Configure the deployment:  
   - **Description**: Add a brief note (e.g., "DataMate FormBuilder").  
   - **Execute as**: Select "Me" (runs under your account).  
   - **Who has access**: Choose "Anyone" (for public access) or "Anyone with a Google account" (restricted to Google users).  
4. Click **Deploy** and copy the generated **Web App URL**.  
5. Share the URL with users to access FormBuilder forms or CSV upload functionality directly in their browsers.  
   - Example use: Deploy `tutorial.html` as a standalone guide or `UploadCSV.html` for remote data uploads.  
6. To update the web app later, click **Deploy > Manage Deployments**, select your deployment, and choose **New Version**.  

**Note**: Web deployment requires `Code.gs` to include functions like `doGet(e)` to serve HTML files (e.g., `return HtmlService.createHtmlOutputFromFile('tutorial')`). Check the repository’s `Code.gs` for implementation details.

---

### Completed Features  
- Spreadsheet-based database system  
- Automatic data logging & organization (up to 12 log fields)  
- Simple UI for data entry  
- FormBuilder with 25+ field types (e.g., Dynamic Tables, Conditional Logic)  
- Batch CSV uploads via `UploadCSV.html`  
- PDF generation from records  
- Tutorial guide via `tutorial.html`  

### Upcoming Features  
- **Multi-user collaboration** – Allow multiple users to access & modify data securely  
- **Automated backups** – Save snapshots of your data for easy recovery  
- **Custom form integrations** – Support for Google Forms & other input sources  
- **Advanced filtering** – Smarter search & filter tools  
- **Collaboration is welcome!** – Need collaborators to develop these features.  

Have feature requests? [Email me!](mailto:datamateapp@gmail.com)  
Visit the website? [Website](https://datamateapp.github.io/)

---

## 💙 Support This Project  

DataMate is free, but if you find it useful, consider supporting development:  

[**Support DataMateApps**](https://datamateapp.github.io/Donate%205%20per%20mo.html)  

Every donation helps keep this project alive and evolving!  

---

## License  
This project is licensed under the **MIT License**. See `LICENSE.txt` for details.  

## Credits  
Developed by **Dan Northway**.  
