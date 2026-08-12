# ==================== PHASE 3 MODULAR AUTOMATION ENGINE ====================
# automation_engine.py - Reusable automation workflows (SYNC Playwright)

from typing import List, Dict, Optional, Tuple
import pandas as pd

from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError

# ==================== AUTHENTICATION HANDLERS ====================

def handle_sso_login(page: Page, username: str, password: str, log_queue: List[str]):
    """
    Handle SSO/IDCS login (without MFA).
    Supports multiple Oracle Cloud SSO form types (SYNC Playwright).
    """
    log_queue.append("🔐 [SSO] Entering credentials...")
    try:
        page.wait_for_selector(
            '#idcs-signin-basic-signin-form-username, input#userid, input[name="username"], input[type="email"]',
            timeout=15000
        )

        # NEW REDWOOD IDCS LOGIN (2024+)
        if page.is_visible('#idcs-signin-basic-signin-form-username'):
            log_queue.append("✓ Detected NEW Redwood IDCS login form")
            page.fill('#idcs-signin-basic-signin-form-username', username)
            page.fill('#idcs-signin-basic-signin-form-password\\|input', password)
            page.click('.oj-button-text:has-text("Sign In")')

        # OLD IDCS
        elif page.is_visible('input#userid'):
            log_queue.append("✓ Detected Oracle IDCS login form")
            page.fill('input#userid', username)
            page.fill('input#password', password)
            page.click('button#btnActive')

        # Standard form
        elif page.is_visible('input[name="username"]'):
            log_queue.append("✓ Detected standard username/password form")
            page.fill('input[name="username"]', username)
            page.fill('input[name="password"]', password)
            page.click('button[type="submit"]')

        # Email-based form
        elif page.is_visible('input[type="email"]'):
            log_queue.append("✓ Detected email-based login form")
            page.fill('input[type="email"]', username)
            page.fill('input[type="password"]', password)
            page.click('button[type="submit"]')

        # Fallback
        else:
            log_queue.append("⚠️ Login form not recognized — using fallback")
            inputs = page.query_selector_all('input')
            if len(inputs) >= 2:
                inputs[0].fill(username)
                inputs[1].fill(password)
                inputs[1].press("Enter")

        log_queue.append("✅ [SSO] Sign In clicked")
        page.wait_for_load_state("networkidle", timeout=30000)

    except PlaywrightTimeoutError:
        log_queue.append("ℹ️ [SSO] No login form detected — assuming already authenticated")


def handle_mfa_login(page: Page, username: str, password: str, log_queue: List[str], mfa_code: Optional[str] = None):
    """
    Handle SSO/IDCS login WITH MFA support (SYNC).
    If mfa_code is None, waits for manual MFA completion (up to 60s).
    """
    log_queue.append("🔐 [MFA] Entering credentials...")
    handle_sso_login(page, username, password, log_queue)

    try:
        el = page.wait_for_selector(
            'input[type="text"][placeholder*="code"], input[aria-label*="verification"], input[name*="otp"]',
            timeout=10000
        )
        if el:
            log_queue.append("🔐 [MFA] Multi-factor authentication detected")
            if mfa_code:
                log_queue.append("🔐 [MFA] Entering provided MFA code...")
                try:
                    el.fill(mfa_code)
                except Exception:
                    page.fill('input[type="text"]', mfa_code)
                if page.is_visible('button[type="submit"]'):
                    page.click('button[type="submit"]')
                else:
                    page.keyboard.press("Enter")
                log_queue.append("✅ [MFA] Code submitted")
            else:
                log_queue.append("⏸️ [MFA] Waiting for manual MFA completion (60 seconds)...")
                log_queue.append("👤 Please complete MFA authentication in the browser")
                page.wait_for_function(
                    """
                    () => {
                        const url = window.location.href.toLowerCase();
                        return !url.includes('mfa') && !url.includes('verify') && !url.includes('challenge');
                    }
                    """,
                    timeout=60000
                )
                log_queue.append("✅ [MFA] Manual MFA completed")
    except PlaywrightTimeoutError:
        log_queue.append("ℹ️ [MFA] No MFA prompt detected")


# ==================== NAVIGATION HANDLERS ====================

def navigate_to_page(page: Page, navigation_path: List[Dict[str, str]], log_queue: List[str]):
    """
    Execute navigation path to reach target Oracle Cloud page.
    """
    log_queue.append("🧭 Starting navigation to target page...")

    for idx, step in enumerate(navigation_path, 1):
        action = step.get("action")
        selector = step.get("selector")

        log_queue.append(f"Step {idx}: {action}")
        try:
            if action in {"click_hamburger", "expand_section"}:
                page.wait_for_selector(selector, timeout=30000)
                page.click(selector)
                page.wait_for_timeout(1000)

            elif action == "click_item":
                page.wait_for_selector(selector, timeout=30000)
                page.click(selector)
                page.wait_for_load_state("networkidle")

            else:
                log_queue.append(f"ℹ️ Unknown action '{action}' (skipped)")

            log_queue.append(f"✓ Step {idx} completed")

        except Exception as e:
            log_queue.append(f"✗ Step {idx} failed: {str(e)}")
            raise


def open_visual_builder(page: Page, log_queue: List[str]) -> Page:
    """
    Open Visual Builder Studio from current page and return the VB Studio Page.
    """
    log_queue.append("🎨 Opening Visual Builder Studio...")

    # Click Avatar
    avatar = 'div.oj-sp-helper-position-relative >> oj-avatar'
    page.wait_for_selector(avatar, timeout=30000)
    page.click(avatar)
    log_queue.append("✓ Avatar clicked")

    # Click "Edit Page in Visual Builder Studio"
    edit_vb_link = 'a#ojSpSimpleUIShellUserProfile_GUMAdministration'
    page.wait_for_selector(edit_vb_link, timeout=30000)

    with page.context.expect_page(timeout=30000) as new_page_info:
        page.click(edit_vb_link)

    vb_page = new_page_info.value
    log_queue.append("⏳ Waiting for Visual Builder Studio tab...")

    vb_page.wait_for_load_state("domcontentloaded", timeout=60000)
    
    # ✅ FIXED: No invalid :has-text() selectors in wait_for_function
    vb_page.wait_for_function(
        """
        () => {
            const url = window.location.href.toLowerCase();
            if (url.includes('identity.oraclecloud.com/oauth2')) return false;
            
            // Check URL patterns
            if (url.includes('/developer/') ||
                url.includes('/vb/') ||
                url.includes('devcsapp') ||
                url.includes('visualbuilder')) {
                return true;
            }
            
            // Check title
            if (document.title.toLowerCase().includes('visual builder')) {
                return true;
            }
            
            // Check for "Configure Fields and Regions" text using valid JavaScript
            const spans = Array.from(document.querySelectorAll('span'));
            if (spans.some(el => el.textContent && el.textContent.includes('Configure Fields and Regions'))) {
                return true;
            }
            
            return false;
        }
        """,
        timeout=90000
    )

    log_queue.append(f"✅ Visual Builder Studio loaded: {vb_page.url}")
    return vb_page


def select_vb_project(vb_page: Page, project_name: str, log_queue: List[str]):
    """
    Select a project in Visual Builder Studio (if the project picker is visible).
    """
    try:
        vb_page.wait_for_selector('input[placeholder="Filter Projects"]', timeout=8000)
        log_queue.append(f"📁 Selecting project: '{project_name}'")

        vb_page.fill('input[placeholder="Filter Projects"]', project_name)
        vb_page.keyboard.press("Enter")
        vb_page.wait_for_timeout(2000)

        # Try exact match first
        project_selector = f'div.oj-listitemlayout-textslots:has-text("{project_name}")'
        try:
            vb_page.wait_for_selector(project_selector, timeout=15000)
            vb_page.click(project_selector)
            log_queue.append(f"✓ Project '{project_name}' selected")
        except PlaywrightTimeoutError:
            # Fallback to partial match
            items = vb_page.query_selector_all('div.oj-listitemlayout-textslots')
            for proj in items:
                text = (proj.text_content() or "").strip()
                if text and project_name.lower() in text.lower():
                    proj.click()
                    log_queue.append(f"✓ Project found and selected: {text}")
                    break

        # Click Select button
        vb_page.wait_for_selector('button:has-text("Select")', timeout=20000)
        vb_page.click('button:has-text("Select")')

        # ✅ FIXED: Wait for editor to load - No invalid selectors in wait_for_function
        vb_page.wait_for_function(
            """
            () => {
                // Check for tree views
                if (document.querySelector('.oj-treeview') || document.querySelector('[role="tree"]')) {
                    return true;
                }
                
                // Check for "Configure Fields and Regions" text using valid JavaScript
                const spans = Array.from(document.querySelectorAll('span'));
                return spans.some(el => el.textContent && el.textContent.includes('Configure Fields and Regions'));
            }
            """,
            timeout=60000
        )
        log_queue.append("✅ Project loaded successfully")

    except PlaywrightTimeoutError:
        log_queue.append("💡 Not on project selection screen — already inside a project")


def navigate_to_vb_page(vb_page: Page, vb_search_term: str, log_queue: List[str]):
    """
    Navigate to specific page in Visual Builder Studio using the page-search control.
    """
    log_queue.append(f"🔍 Searching for page: '{vb_search_term}'")

    # List of possible page names for the search dropdown button
    possible_button_names = [
        "Homepage - Search",
        "Approval History - View",
        "Delivery Billing - Edit",
        "Document History - Public View",
        "Document History - View",
        "Edit Change Order",
        "Extended Contract Lines - Search",
        "Marketplace Items - Search",
        "Multiple Lines - Edit",
        "Noncatalog Request - Edit",
        "Product Details - View",
        "Requisition - Edit",
        "Requisition Details - Public View",
        "Requisition Details - View",
        "Requisition Lifecycle - View",
        "Requisition Line - Edit",
        "Requisitions - Manage",
        "Self Service Procurement Home - Search",
        "Shopping Cart - View",
        "Shopping Error - View",
        "Shopping List - Edit",
        "Shopping List - View",
        "Shopping Lists - Search"
    ]

    # Try to find and click the dropdown button
    dropdown_button = None
    button_found = None
    
    for button_name in possible_button_names:
        try:
            dropdown_button = vb_page.get_by_role("button", name=button_name)
            dropdown_button.wait_for(state="visible", timeout=2000)
            button_found = button_name
            log_queue.append(f"✓ Found dropdown button: '{button_name}'")
            break
        except PlaywrightTimeoutError:
            continue
    
    if not dropdown_button or not button_found:
        log_queue.append("⚠️ Could not find page dropdown button by name, trying fallback...")
        # Fallback: Try to find button by aria-label containing "Search"
        try:
            dropdown_button = vb_page.locator('button[aria-label*="Search"]').first
            dropdown_button.wait_for(state="visible", timeout=5000)
            log_queue.append("✓ Found dropdown button using fallback selector")
        except:
            raise Exception("Could not find page search dropdown button")
    
    dropdown_button.click()
    log_queue.append("✓ Search dropdown opened")

    # Enter search term
    search_input = vb_page.locator('input[aria-label="Filter Search"]')
    search_input.wait_for(state="visible", timeout=20000)
    search_input.fill(vb_search_term)
    search_input.press("Enter")
    log_queue.append("✓ Search submitted")

    # Select the result
    page_item = vb_page.get_by_text(vb_search_term)
    page_item.wait_for(state="visible", timeout=20000)
    page_item.click()
    log_queue.append(f"✅ Page '{vb_search_term}' opened")
    vb_page.wait_for_timeout(2000)


# ==================== VB STUDIO URL CAPTURE ====================

def capture_vb_urls(vb_page: Page, log_queue: List[str]) -> dict:
    """
    Capture both Editor URL and Preview URL from VB Studio.
    Returns dict with 'editor_url' and 'preview_url'.
    """
    log_queue.append("📸 Capturing VB Studio URLs...")
    urls: Dict[str, str] = {}

    try:
        # Editor URL
        editor_url = vb_page.url
        urls["editor_url"] = editor_url
        log_queue.append(f"✓ Editor URL captured: {editor_url[:80]}...")

        # Preview URL
        log_queue.append("🖱️ Clicking Preview button to capture preview URL...")
        button_selector = 'button[aria-label="Preview"]'

        vb_page.wait_for_selector(button_selector, state="visible", timeout=20000)
        vb_page.eval_on_selector(button_selector, "el => el.scrollIntoView({block: 'center'})")

        with vb_page.context.expect_page(timeout=20000) as new_page_info:
            vb_page.click(button_selector)

        preview_page = new_page_info.value
        preview_page.wait_for_load_state("networkidle", timeout=30000)
        preview_url = preview_page.url
        urls["preview_url"] = preview_url
        log_queue.append(f"✓ Preview URL captured: {preview_url[:80]}...")

        # Close preview tab (optional)
        preview_page.close()
        log_queue.append("✓ Preview tab closed")

        log_queue.append("✅ Both URLs captured successfully")
        return urls

    except Exception as e:
        log_queue.append(f"⚠️ Failed to capture URLs: {str(e)}")
        return urls  # may contain only editor_url


# ==================== VB STUDIO OPERATIONS ====================

def create_customization_rule(vb_page: Page, rule_name: str, log_queue: List[str]):
    """
    Create a new customization rule in VB Studio (Advanced mode will be enabled later).
    """
    log_queue.append("📝 Creating customization rule...")

    # Click "Configure Fields and Regions"
    config_button = 'span:has-text("Configure Fields and Regions")'
    vb_page.wait_for_selector(config_button, timeout=60000)
    vb_page.click(config_button)
    log_queue.append("✓ 'Configure Fields and Regions' clicked")

    # Click "Create Rule"
    create_rule_btn = 'oj-button[data-vb-id*="createBtn"]'
    vb_page.wait_for_selector(create_rule_btn, timeout=60000)
    vb_page.click(create_rule_btn)
    log_queue.append("✓ 'Create Rule' clicked")

    # Fill rule name
    rule_input_selector = 'input.oj-text-field-input[aria-required="true"]'
    vb_page.wait_for_selector(rule_input_selector, timeout=30000)
    input_element = vb_page.query_selector(rule_input_selector)
    input_element.scroll_into_view_if_needed()
    input_element.focus()
    input_element.evaluate('(el) => el.value = ""')
    input_element.type(rule_name, delay=100)
    input_element.press("Enter")
    log_queue.append(f"✅ Rule created: '{rule_name}'")
    vb_page.wait_for_timeout(2000)


def activate_advanced_mode(vb_page: Page, log_queue: List[str]):
    """
    Activate Advanced mode in VB Studio customization.
    """
    log_queue.append("🔧 Activating Advanced mode...")
    advanced_input_selector = 'input[type="radio"][value="Advanced"]'
    vb_page.wait_for_selector(advanced_input_selector, timeout=30000)
    is_checked = vb_page.is_checked(advanced_input_selector)

    if not is_checked:
        advanced_button = 'span.oj-button-text:has-text("Advanced")'
        vb_page.wait_for_selector(advanced_button, timeout=20000)
        vb_page.click(advanced_button)
        log_queue.append("✓ Advanced mode activated")
        vb_page.wait_for_timeout(2000)
    else:
        log_queue.append("✓ Already in Advanced mode")


def open_metadata_file(vb_page: Page, vb_page_file: str, log_queue: List[str]):
    """
    Open the metadata-rules-x.json file for a specific VB page in the Source tab.
    """
    log_queue.append(f"📄 Opening metadata file for '{vb_page_file}'...")

    # Click Source tab
    source_tab_selector = 'li#appNavSource a[role="tab"]'
    vb_page.wait_for_selector(source_tab_selector, timeout=30000)
    is_selected = vb_page.get_attribute(source_tab_selector, "aria-selected")
    if is_selected != "true":
        vb_page.click(source_tab_selector)
        vb_page.wait_for_timeout(2000)
    log_queue.append("✓ Source tab activated")

    # Search for metadata file
    search_input_selector = 'input[aria-label="Filter Search"]'
    vb_page.wait_for_selector(search_input_selector, timeout=30000)
    vb_page.fill(search_input_selector, vb_page_file)
    vb_page.locator(search_input_selector).press("Enter")
    vb_page.wait_for_timeout(2000)
    log_queue.append("✓ Search completed")

    # Click metadata file
    metadata_file_selector = 'text="metadata-rules-x.json"'
    vb_page.wait_for_selector(metadata_file_selector, timeout=40000)
    vb_page.click(metadata_file_selector)
    vb_page.wait_for_timeout(2000)
    log_queue.append("✓ Metadata file opened")

    # Wait for Monaco editor
    vb_page.wait_for_function(
        """
        () => {
            const models = window.monaco?.editor?.getModels();
            return models && models.length > 0 && models[0].getValue().length > 0;
        }
        """,
        timeout=60000
    )
    log_queue.append("✅ Monaco editor loaded")


# ==================== EXCEL PROCESSING ====================

def parse_personalization(p_str: str) -> Optional[Tuple[str, bool]]:
    """Parse personalization string from Excel."""
    s = str(p_str).strip()
    if s in ("Element re-ordered(mds:move)", ""):
        return None
    if "=" not in s:
        return None
    k, v = s.split("=", 1)
    k, v = k.strip().lower(), v.strip().lower()

    if k == "rendered":
        return ("hidden", v == "false")
    elif k == "readonly":
        return ("readonly", v == "true")
    elif k in ("required", "showrequired"):
        return ("required", v == "true")

    return None


def process_excel_mapping(
    excel_file: str,
    composite_mapping: Dict[Tuple[str, str], str],
    log_queue: List[str]
) -> Tuple[Dict, List[Tuple], List[Tuple]]:
    """
    Process Excel file and generate field mappings.

    Returns:
        fields: Dictionary of field personalizations
        mapped_records: List of successfully mapped records
        unmapped_records: List of unmapped records
    """
    log_queue.append(f"📊 Processing Excel: {excel_file}")

    df = pd.read_excel(excel_file, dtype=str)  # ensure openpyxl installed for .xlsx
    required_cols = {"File", "Field", "Display Name", "Personalization"}
    if not required_cols <= set(df.columns):
        missing = required_cols - set(df.columns)
        raise ValueError(f"Missing columns: {missing}")

    fields: Dict[str, Dict] = {}
    mapped_records: List[Tuple] = []
    unmapped_records: List[Tuple] = []

    skip_display_names = {
        "Column", "Panel Form Layout", "Region", "Delivery", "Billing", "Tax",
        "NotesandAttachments", "Source", "Popup", "Output Text",
        "Panel Group Layout", "DocumentAttachments panelHeader"
    }

    for _, r in df.iterrows():
        file = str(r["File"]).strip()
        field = str(r["Field"]).strip()
        disp = str(r["Display Name"]).strip()
        pers = r["Personalization"]

        # Skip header-like noise rows
        if file.lower() == "file":
            continue

        key = (file, field)
        new_key = composite_mapping.get(key)

        if new_key is None:
            if disp not in skip_display_names:
                unmapped_records.append((file, field, disp, str(pers)))
            continue

        prop = parse_personalization(pers)
        if not prop:
            continue

        p_key, p_val = prop
        if new_key not in fields:
            fields[new_key] = {}
        fields[new_key][p_key] = {"value": p_val}
        mapped_records.append((file, field, disp, new_key))

    log_queue.append(f"✓ Mapped: {len(mapped_records)}, Unmapped: {len(unmapped_records)}")
    return fields, mapped_records, unmapped_records


def inject_personalization_to_monaco(vb_page: Page, fields: Dict, log_queue: List[str]) -> bool:
    """
    Inject personalization fields into Monaco editor (within the current metadata-rules-x.json).
    """
    log_queue.append("💉 Injecting personalization into Monaco editor...")

    result = vb_page.evaluate(
        """(newFields) => {
            try {
                const models = window.monaco?.editor?.getModels?.();
                if (!models || models.length === 0) 
                    return { success: false, error: "No editor model" };

                const model = models[0];
                const content = model.getValue();
                const data = JSON.parse(content);

                if (data.addMetadataRules?.[0]?.overlay) {
                    data.addMetadataRules[0].overlay.fields = newFields;
                    model.setValue(JSON.stringify(data, null, 2));
                    return { success: true };
                } else {
                    return { success: false, error: "Invalid JSON structure" };
                }
            } catch (e) {
                return { success: false, error: e.message };
            }
        }""",
        fields
    )

    if result.get("success"):
        log_queue.append("✅ Personalization injected successfully!")
        return True
    else:
        log_queue.append(f"❌ Injection failed: {result.get('error')}")
        return False