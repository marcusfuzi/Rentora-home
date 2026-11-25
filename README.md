# 🏠 Rentora - Smart Rental Management System

**Rentora** is a modern, full-stack web application designed to eliminate intermediary brokers and streamline the process of finding and managing rental accommodations. Built with Python and Django, Rentora directly connects property owners and prospective tenants with zero brokerage fees, transparent pricing, rich photographic previews, and an intelligent multi-attribute search and filtering engine.

---

## 👥 Project Team & Academic Credentials

* **Degree:** Bachelor of Technology (B.Tech) in Computer Science & Engineering
* **Institution:** Department of Computer Science and Engineering, Dr. APJ. Abdul Kalam Institute of Technology, Tanakpur (Champawat, Uttarakhand)
* **University:** Veer Madho Singh Bhandari Uttarakhand Technical University (VMSBUTU), Dehradun
* **Academic Session:** 2026 – 2027
* **Project Mentor:** Assistant Prof. **Mr. Divisht Jaiswal** (Dept. of CSE)
* **Head of Department:** **Mr. Naresh Kumar** (Dept. of CSE)

### Team Members:
1. **Tushar Singh Bisht** (Roll No: `2316201064`) - *Lead Full-Stack Architecture & Database*
2. **Ayush Kumar Jha** (Roll No: `2316201019`) - *Frontend UI/UX & Responsive Design*
3. **Abhishek Kumar Arya** (Roll No: `2316201004`) - *Security, RBAC & Media Pipeline*

---

## 🚀 Key Features

* **Zero Brokerage Marketplace:** Direct peer-to-peer connection between landlords and prospective tenants without third-party commissions.
* **Smart Visitor Search Console & Filter Sidebar:**
  * 📍 **Location / City Search:** Fast text querying for specific cities and neighborhoods.
  * 🏢 **Property Type Selector:** Apartment, Independent House, Villa, Studio, Commercial.
  * 🛏️ **BHK Configuration:** 1 BHK, 2 BHK, 3 BHK, 4+ BHK pills.
  * 💰 **Dual Budget Range & Interactive Slider:** Filter minimum to maximum rent (₹5,000 to ₹45,000+).
  * 🛋️ **Furnishing Status:** Fully-Furnished, Semi-Furnished, Unfurnished.
  * 👥 **Preferred Tenant:** Family Only, Bachelors / Students Friendly, Single Professional, Any.
  * 🚗 **Parking Availability:** Dedicated Car Parking, Two-Wheeler Only.
  * ☑️ **Must-Have Amenities:** Wi-Fi, 24x7 Water Supply, Power Backup, CCTV/Security, Lift.
  * 🔄 **Smart Sorting:** Rent (Low to High / High to Low), Star Ratings, Recently Added.
  * ⚡ **One-Click Quick Presets:** Dedicated quick buttons for *College Students*, *Working Professionals*, and *Families*.
* **Role-Based Portals (RBAC):** Dedicated, secure dashboards for Property Owners and Visitors/Tenants.
* **Media & Photo Engine:** High-resolution image upload pipeline with responsive card previews and detail modals.
* **Security Hardening:** Automatic CSRF validation, PBKDF2 with SHA-256 password cryptography, and SQL injection immunity via Django ORM.

---

## 🛠️ Technology Stack

* **Backend Framework:** Django 3.2 LTS (Python 3.10)
* **Architecture:** Model-View-Template (MVT)
* **Database:** SQLite (development) / MySQL & PostgreSQL compatible via Django ORM
* **Frontend:** HTML5, CSS3, JavaScript (ES6+), Bootstrap, FontAwesome
* **Image Processing:** Pillow (PIL)

---

## 💻 Installation & Quickstart

1. **Clone the repository:**
   ```bash
   git clone https://github.com/marcusfuzi/Rentora-home.git
   cd Rentora-home/rental1
   ```

2. **Create and activate a virtual environment:**
   ```bash
   python -m venv .venv
   # Windows:
   .\.venv\Scripts\activate
   # Linux / macOS:
   source .venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install django pillow
   ```

4. **Apply database migrations:**
   ```bash
   python manage.py migrate
   ```

5. **Start the development server:**
   ```bash
   python manage.py runserver
   ```
   Open your browser and navigate to `http://127.0.0.1:8000/`.

---

## 📄 License
This project is developed for educational and academic purposes under the B.Tech CSE curriculum at Dr. APJ Abdul Kalam Institute of Technology, Tanakpur.
