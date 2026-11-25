import os
from playwright.sync_api import sync_playwright

OUT_DIR = r"C:\Users\bisht\.gemini\antigravity\scratch\Rental_Management_System\report_figures"
os.makedirs(OUT_DIR, exist_ok=True)
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def capture():
    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=EDGE_PATH, headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 850})
        page = context.new_page()

        # 1. Homepage
        print("Capturing 1. Homepage...")
        page.goto("http://127.0.0.1:8000/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_1_home.png"))

        # 2. Login Page
        print("Capturing 2. Login Page...")
        page.goto("http://127.0.0.1:8000/login/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_2_login.png"))

        # 3. Owner Register Page
        print("Capturing 3. Owner Register...")
        page.goto("http://127.0.0.1:8000/register/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_3_register_owner.png"))

        # 4. Visitor Register Page
        print("Capturing 4. Visitor Register...")
        page.goto("http://127.0.0.1:8000/registerVisitor/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_4_register_visitor.png"))

        # 5. Log in as Owner
        print("Logging in as john_owner...")
        page.goto("http://127.0.0.1:8000/login/", wait_until="networkidle")
        page.fill("input[name='username']", "john_owner")
        page.fill("input[name='password']", "pass1234")
        page.click("button[type='submit'], input[type='submit']")
        page.wait_for_load_state("networkidle")

        print("Capturing 5. Owner Dashboard...")
        page.goto("http://127.0.0.1:8000/login/ownerDashboard/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_5_owner_dashboard.png"))

        print("Capturing 6. Add Property Form...")
        page.goto("http://127.0.0.1:8000/login/addProperty/", wait_until="networkidle")
        page.screenshot(path=os.path.join(OUT_DIR, "screen_6_add_property.png"))

        # Clear cookies / context for tenant
        context.close()
        context2 = browser.new_context(viewport={"width": 1280, "height": 850})
        page2 = context2.new_page()

        # 7. Log in as Tenant
        print("Logging in as alice_tenant...")
        page2.goto("http://127.0.0.1:8000/login/", wait_until="networkidle")
        page2.fill("input[name='username']", "alice_tenant")
        page2.fill("input[name='password']", "pass1234")
        page2.click("button[type='submit'], input[type='submit']")
        page2.wait_for_load_state("networkidle")

        print("Capturing 7. Visitor Dashboard...")
        page2.goto("http://127.0.0.1:8000/dashboard/", wait_until="networkidle")
        page2.screenshot(path=os.path.join(OUT_DIR, "screen_7_visitor_dashboard.png"))

        print("Capturing 8. Property Detail View...")
        page2.goto("http://127.0.0.1:8000/viewProp/1/", wait_until="networkidle")
        page2.screenshot(path=os.path.join(OUT_DIR, "screen_8_property_detail.png"))

        browser.close()
        print("All 8 UI screenshots captured successfully!")

if __name__ == "__main__":
    capture()
