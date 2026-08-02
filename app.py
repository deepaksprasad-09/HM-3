from flask import Flask, render_template, request, jsonify, Response
import openpyxl, threading, time, json, queue, os, traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True)

log_queue = queue.Queue()
bot_status = {
    'running': False, 'total': 0, 'done': 0,
    'errors': 0, 'current_row': {}, 'current_idx': 0, 'step': ''
}
WAIT_SEC = 5

# ── Helpers ───────────────────────────────────────────────────────────────────
def log(type_, msg):
    entry = {'type': type_, 'msg': msg, 'time': time.strftime('%H:%M:%S')}
    log_queue.put(json.dumps(entry))
    print('[' + type_.upper() + '] ' + msg)

def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def safe_fill(driver, wait, locator, value):
    el = wait.until(EC.presence_of_element_located(locator))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    el.clear()
    el.send_keys(str(value))

def wait_load(driver, min_els=3, max_wait=8):
    for _ in range(max_wait):
        time.sleep(0.3)
        if len(driver.find_elements(By.XPATH, "//*[string-length(normalize-space(text())) > 2]")) >= min_els:
            return True
    return False

# ── Read Excel ────────────────────────────────────────────────────────────────
def read_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    raw = [str(h).strip() if h else '' for h in rows[0]]
    # Make duplicate headers unique
    seen = {}
    headers = []
    for h in raw:
        if h in seen:
            seen[h] += 1
            headers.append(h + '_' + str(seen[h]))
        else:
            seen[h] = 0
            headers.append(h)
    records = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        rec = {headers[i]: (str(row[i]).strip() if i < len(row) and row[i] is not None else '')
               for i in range(len(headers))}
        if 'CONTRACT_NO' not in rec and 'CONTRACT_NO_1' in rec:
            rec['CONTRACT_NO'] = rec['CONTRACT_NO_1']
        records.append(rec)
    return records

# ── Bot ───────────────────────────────────────────────────────────────────────
def run_bot(filepath):
    global bot_status
    try:
        records = read_excel(filepath)
        if not records:
            log('error', 'No records in Excel!'); bot_status['running'] = False; return

        bot_status.update({'total': len(records), 'done': 0, 'errors': 0})
        log('info', 'Starting bot for ' + str(len(records)) + ' records...')

        opts = webdriver.ChromeOptions()
        opts.add_argument('--start-maximized')
        opts.add_argument('--disable-blink-features=AutomationControlled')
        opts.add_experimental_option('excludeSwitches', ['enable-automation'])
        opts.add_experimental_option('useAutomationExtension', False)

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
        driver.execute_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
        wait = WebDriverWait(driver, WAIT_SEC)
        logged_in_user = None

        for idx, row in enumerate(records, 1):
            username = row.get('id', '')
            password = row.get('pass', '')
            contract = row.get('CONTRACT_NO', '')
            bot_status.update({'current_idx': idx, 'current_row': row})

            log('info', '━━━ Row ' + str(idx) + '/' + str(len(records)) + ' | Contract: ' + contract + ' ━━━')

            # LOGIN
            if logged_in_user != username:
                bot_status['step'] = 'login'
                log('step', 'Logging in as ' + username + '...')
                driver.get('https://ereceipt.homecredit.co.in/cuppaymentfieldexecutiveapp/UI/login')
                wait_load(driver, 2, 6); time.sleep(0.2)
                for loc in [(By.NAME,'user'),(By.NAME,'username'),(By.XPATH,"//input[@type='text']")]:
                    try: safe_fill(driver,wait,loc,username); break
                    except: continue
                for loc in [(By.NAME,'password'),(By.XPATH,"//input[@type='password']")]:
                    try: safe_fill(driver,wait,loc,password); break
                    except: continue
                for loc in [(By.XPATH,"//input[@type='submit']"),(By.XPATH,"//button[@type='submit']"),(By.XPATH,"//button")]:
                    try: el=wait.until(EC.element_to_be_clickable(loc)); js_click(driver,el); break
                    except: continue
                for _ in range(12):
                    time.sleep(0.3)
                    if 'login' not in driver.current_url.lower(): break
                log('success', 'Logged in → ' + driver.current_url)

                # PAYMENT COLLECTION
                bot_status['step'] = 'payment_collection'
                log('step', 'Selecting Payment Collection...')
                wait_load(driver, 3, 3); time.sleep(0.5)
                clicked = False
                for el in driver.find_elements(By.XPATH, "//*"):
                    try:
                        txt = el.text.strip().replace(' ','').replace('\n','').lower()
                        if 'paymentcollection' in txt:
                            js_click(driver, el); time.sleep(0.4)
                            if 'searchcontract' in driver.current_url:
                                clicked = True; log('success','Payment Collection selected!'); break
                    except: pass
                if not clicked:
                    log('warning','Auto-click failed. Please click Payment Collection manually (15s)...')
                    for _ in range(15):
                        time.sleep(1)
                        if 'searchcontract' in driver.current_url:
                            log('success','Manual click detected!'); break
                logged_in_user = username
            else:
                wait_load(driver, 3, 3)

            # SEARCH CONTRACT
            bot_status['step'] = 'search'
            log('step', 'Searching contract: ' + contract)
            wait_load(driver, 3, 4); time.sleep(0.1)
            inputs = driver.find_elements(By.XPATH, "//input")
            filled = 0
            for inp in inputs:
                try:
                    tp = (inp.get_attribute('type') or 'text').lower()
                    if tp in ('submit','button','hidden','checkbox','radio'): continue
                    if inp.is_displayed():
                        inp.clear(); inp.send_keys(str(contract)); filled += 1
                        if filled >= 2: break
                except: pass
            for loc in [(By.XPATH,"//button[contains(text(),'Search')]"),(By.XPATH,"//input[@type='submit']"),(By.XPATH,"//button")]:
                try: el=wait.until(EC.element_to_be_clickable(loc)); js_click(driver,el); log('success','Contract found!'); break
                except: continue
            wait_load(driver, 4, 4); time.sleep(0.1)

            # FILL FORM
            bot_status['step'] = 'fill'
            log('step', 'Filling form...')
            amt = row.get('collected_amount','0')
            for loc in [(By.XPATH,"//input[@placeholder='Enter Amount Collected']"),(By.XPATH,"//input[contains(@placeholder,'Amount')]"),(By.NAME,'collectedAmount')]:
                try: safe_fill(driver,wait,loc,amt); log('step','  ✔ Amount='+amt); break
                except: continue
            pname = row.get("payer's_name",'')
            for loc in [(By.XPATH,"//input[@placeholder=\"Enter Payer's Name\"]"),(By.XPATH,"//input[contains(@placeholder,'Payer')]"),(By.NAME,'name')]:
                try: safe_fill(driver,wait,loc,pname); log('step','  ✔ Payer='+pname); break
                except: continue
            mobile = row.get("payer's_mobile_no",'')
            for loc in [(By.XPATH,"//input[@placeholder='Enter Mobile Number']"),(By.XPATH,"//input[contains(@placeholder,'Mobile')]"),(By.NAME,'mobile')]:
                try: safe_fill(driver,wait,loc,mobile); log('step','  ✔ Mobile='+mobile); break
                except: continue
            try:
                for sel in driver.find_elements(By.TAG_NAME,'select'):
                    opts=[o.text.strip() for o in sel.find_elements(By.TAG_NAME,'option')]
                    if any('Not Collected' in o for o in opts):
                        Select(sel).select_by_visible_text('Not Collected'); log('step','  ✔ Status=Not Collected'); time.sleep(1); break
            except: pass
            ptp = False
            for _ in range(5):
                try:
                    for sel in driver.find_elements(By.TAG_NAME,'select'):
                        opts=[o.text.strip() for o in sel.find_elements(By.TAG_NAME,'option')]
                        if any('PTP' in o for o in opts):
                            Select(sel).select_by_visible_text('PTP - PROMISE TO PAY'); log('step','  ✔ PTP selected'); ptp=True; break
                    if ptp: break
                except: pass
                time.sleep(0.5)
            addr = row.get('new_address','NA')
            for loc in [(By.XPATH,"//input[@placeholder='Enter the new address...']"),(By.NAME,'address')]:
                try: safe_fill(driver,wait,loc,addr); log('step','  ✔ Address='+addr); break
                except: continue
            comment = next((v for k,v in row.items() if 'CM PTP' in k.upper() or 'COMMENT' in k.upper()),'')
            for loc in [(By.XPATH,"//textarea[@placeholder='Please enter your comments...']"),(By.XPATH,"//textarea"),(By.NAME,'comments')]:
                try: safe_fill(driver,wait,loc,comment); log('step','  ✔ Comments='+comment); break
                except: continue

            # SUBMIT
            bot_status['step'] = 'submit'
            log('step','Clicking Submit...')
            for loc in [(By.XPATH,"//button[contains(text(),'Submit')]"),(By.XPATH,"//input[@value='Submit']"),(By.XPATH,"//button[@type='submit']")]:
                try: el=wait.until(EC.element_to_be_clickable(loc)); js_click(driver,el); log('success','Submitted!'); break
                except: continue
            time.sleep(0.4)

            # CONFIRM
            bot_status['step'] = 'confirm'
            log('step','Looking for Confirm...')
            confirmed = False
            for xpath in [
                "//button[contains(text(),'Confirm')]",
                "//button[normalize-space(text())='Confirm']",
                "//input[@value='Confirm']",
                "//div[contains(@class,'modal')]//button[1]",
                "//button[not(contains(text(),'Cancel')) and not(contains(text(),'Pay')) and not(contains(text(),'Submit')) and not(contains(text(),'Search'))]"
            ]:
                if confirmed: break
                try:
                    for el in driver.find_elements(By.XPATH, xpath):
                        if el.is_displayed():
                            js_click(driver,el); time.sleep(0.2); confirmed=True
                            log('success','Confirmed! ✅'); break
                except: pass
            if not confirmed:
                log('warning','Waiting 8s for manual Confirm click...')
                for _ in range(8):
                    time.sleep(0.8)
                    btns = driver.find_elements(By.XPATH,"//button[contains(text(),'Confirm')]")
                    if not any(b.is_displayed() for b in btns): break

            bot_status['done'] = idx
            time.sleep(0.1)

        log('success','🎉 ALL ' + str(len(records)) + ' ROWS DONE!')
    except Exception as e:
        bot_status['errors'] += 1
        log('error', 'Bot error: ' + str(e))
        log('error', traceback.format_exc())
    finally:
        try: driver.quit()
        except: pass
        bot_status['running'] = False
        bot_status['step'] = ''
        log('info', 'Browser closed.')

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file part'}), 400
        f = request.files['file']
        if not f or f.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        if not (f.filename.lower().endswith('.xlsx') or f.filename.lower().endswith('.xls')):
            return jsonify({'error': 'Please upload .xlsx or .xls'}), 400
        fname = secure_filename(f.filename)
        path = os.path.join(UPLOAD_FOLDER, fname)
        f.save(path)
        records = read_excel(path)
        return jsonify({'success': True, 'rows': len(records), 'path': path, 'filename': fname})
    except Exception as e:
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500

@app.route('/start', methods=['POST'])
def start():
    global bot_status
    if bot_status['running']:
        return jsonify({'error': 'Already running'}), 400
    data = request.get_json(force=True)
    filepath = data.get('path','')
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found: ' + filepath}), 400
    while not log_queue.empty():
        log_queue.get()
    bot_status['running'] = True
    threading.Thread(target=run_bot, args=(filepath,), daemon=True).start()
    return jsonify({'success': True})

@app.route('/status')
def status():
    return jsonify(bot_status)

@app.route('/stream')
def stream():
    def generate():
        while True:
            try:
                msg = log_queue.get(timeout=25)
                yield 'data: ' + msg + '\n\n'
            except queue.Empty:
                yield 'data: ' + json.dumps({'type':'ping','msg':'','time':''}) + '\n\n'
    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control':'no-cache','X-Accel-Buffering':'no'})

if __name__ == '__main__':
    print('\n' + '='*50)
    print('  Home Credit Auto Entry Tool')
    print('  Open browser: http://127.0.0.1:5000')
    print('='*50 + '\n')
    app.run(debug=False, port=5000, threaded=True, host='0.0.0.0')
