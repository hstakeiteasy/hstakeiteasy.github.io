# HsTakeItEasy - Tuition Management Portal 🎓

**HsTakeItEasy** is a comprehensive, mobile-first web ecosystem built specifically for **Hasibul Sir's English Tuition** (Classes 5 to 12) located in Samudragarh. It provides a modern, seamless digital experience for both the administration and the students.

## 🌟 Key Features

### 👨‍🎓 Student Dashboard (`student.html`)
* **Real-time Statistics:** Instant visibility into overall attendance percentages and recent exam scores.
* **Smart Billing System:** Track total pending due amounts and view detailed monthly bill histories (Paid, Partial, Unpaid).
* **Academic Tracking:** Toggle between Marks and Attendance history to monitor academic progress.
* **Quick Links Hub:** Direct access to Community Groups, Google Drive Materials, Resources, Upcoming Tests, and Important Notices.
* **Interactive UI:** Features a dynamic 30-second smooth-loading screen with background data fetching, an auto-sliding image carousel, and a daily motivational quote.

### 👨‍🏫 Admin Portal (`admin.html`)
* **Complete Management:** Tools to add/edit students, manage batches, record daily attendance, enter test marks, and generate monthly fee bills.
* **Status Tracking:** Easily monitor who has paid fees and who has outstanding dues.

### 🌍 Landing & Public Pages
* **`index.html`**: A highly converting, SEO-optimized landing page featuring Hasibul Sir's services, guaranteed results, and quick access to the portal.
* **`app.html`**: A dedicated download page for the official Android APK.
* **`contact.html` & `privacy.html`**: Professional contact information and data privacy guidelines.
* **`login.html`**: Secure, role-based login system (Admin/Student) with automatic redirection.

## 🛠️ Technology Stack

* **Frontend:** HTML5, CSS3, JavaScript (Vanilla)
* **Styling:** Tailwind CSS (via CDN) for highly responsive, modern, and mobile-friendly UI.
* **Icons:** FontAwesome 6 (No emojis used for better cross-platform consistency).
* **Backend API:** Google Apps Script (`Code.gs`).
* **Database:** Google Sheets (Students, Batches, Bills, Attendance, Marks).
* **Hosting:** Designed to be easily hosted on GitHub Pages or any static web host.

## 📂 Project Structure

```text
├── index.html       # Main landing page for the tuition center
├── login.html       # Role-based login portal (Student/Admin)
├── student.html     # Comprehensive student dashboard
├── admin.html       # Administrative management portal
├── app.html         # Official Android App (APK) download page
├── contact.html     # Contact information and support
├── privacy.html     # Data privacy policy
├── test.html        # Upcoming tests and schedules placeholder
└── README.md        # Project documentation
```

## 🚀 How It Works

1. **Database:** All data is securely stored in a centralized Google Spreadsheet.
2. **API Layer:** A Google Apps Script (`Code.gs`) reads and writes to the spreadsheet and exposes a secure API endpoint using `doGet()` and `doPost()`.
3. **Frontend Integration:** The HTML files use JavaScript `fetch()` API calls (with cache-busting `_t` parameters to ensure live data) to communicate with the Google Apps Script URL.
4. **Authentication:** Uses secure token-based generation upon login, storing the session locally for a smooth UX.

## 📞 Contact & Support

* **Tuition Location:** Samudragarh, Purba Bardhaman, West Bengal - 713519

* **Email:** hstakeiteasy007@gmail.com

---
*Designed & Developed with ❤️ by **Sumit Haldar** (AI Engineer)*  
*[hello@sumitaidev.com](mailto:hello@sumitaidev.com)*
