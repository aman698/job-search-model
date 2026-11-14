# Embedded Job Checker (Noida / Gurgaon)

Excel list ki har company ke careers portal pe embedded / firmware jobs dhoondta hai.

## Setup (ek baar)
```
pip install -r requirements.txt
```
`Embedded_Companies_Noida_Gurgaon.xlsx` isi folder mein rakho.

## Run
```
python job_checker.py                 # ek baar
python job_checker.py --every 60      # har 60 min mein (laptop on rakho)
python job_checker.py --any-location  # India ke bahar / Bengaluru bhi dikhe
```
Result: `results/jobs_YYYYMMDD_HHMM.xlsx`
- **Jobs** sheet: title, location, posted date, direct link. Pichli run ke baad aayi jobs pe **NEW**.
- **Portal status** sheet: kis company ka portal mila, kis pe error aaya.

## Kaise kaam karta hai
1. Company ki website se careers page dhoondta hai (`/careers`, `careers.<domain>`, etc.)
2. Portal pehchanta hai: Workday, Greenhouse, Lever, SmartRecruiters, Oracle HCM, Eightfold.
   Inke public job APIs se "embedded" aur "firmware" search karta hai.
3. Simple HTML careers page ho to page ke links mein embedded/firmware wale titles uthata hai.
4. Title filter: embedded, firmware, BSP, AUTOSAR, RTOS, device driver, bootloader, MCU, ECU, J1939, UDS...
   Location filter: Noida, Gurgaon/Gurugram, Manesar, Delhi, NCR, Faridabad, Ghaziabad.
   Dono `job_checker.py` ke top pe `ROLE_RE` / `LOC_RE` mein badal sakte ho.
5. Mila hua portal `.portal_cache.json` mein save hota hai, agli baar fast chalta hai.

## Jab kisi company ka result na aaye
"Portal status" sheet mein NOT FOUND / ERROR / generic dekho, us company ka job-search page
browser mein kholo aur URL `portals.yaml` mein daal do. Phir dobara run karo.

JS se load hone wale pages (SuccessFactors, Taleo, custom React sites) ke liye:
```
pip install playwright
python -m playwright install chromium
python job_checker.py --render
```

## Limits (seedhi baat)
- LinkedIn / Naukri scrape nahi karta: unki terms allow nahi karti, aur wo block bhi kar dete hain.
  Unke liye unke apne job alerts lagao.
- Kuch portals bots ko block karte hain (403 / captcha). Wo "Portal status" mein ERROR dikhenge,
  unhe manually check karo.
- Workday ki location kabhi "2 Locations" aati hai. Aisi jobs "Location clear nahi" note ke saath
  rakhi jaati hain taaki miss na ho.
- Pehli run mein discovery ki wajah se 3-8 min lag sakte hain.
