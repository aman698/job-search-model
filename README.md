# Embedded Job Checker v2 (Delhi NCR: Noida / Greater Noida / Gurugram / Manesar)

Excel list ki har company ke **official careers portal** pe embedded / firmware jobs dhoondta hai
aur ek Excel deta hai, jisme sab hota hai: konsi company mein vacancy hai, link, experience, match.

## Setup (sirf ek baar)
```
pip install -r requirements.txt
```
`Embedded_Companies_Noida_Gurgaon.xlsx` aur `portals.yaml` isi folder mein rakho.

## Run
```
python job_checker.py                    # ek baar (5-10 min)
python job_checker.py --every 360        # har 6 ghante mein dobara (laptop on rakho)
python job_checker.py --only "Maruti"    # sirf ek company check karo
python job_checker.py --any-location     # NCR filter band (Bengaluru/Pune wali jobs bhi)
python job_checker.py --selftest         # bina internet ke parser test
```

Result: `results/Embedded_Jobs_YYYY-MM-DD_HHMM.xlsx`, 3 sheets:

| Sheet | Kya hai |
|---|---|
| **Jobs** | Company, City, Job title, Location, Experience, Posted date, Key skills, Match (High/Med/Low), Why, Direct apply link, **NEW** (pichli run ke baad aayi) |
| **Portal Status** | Excel ki HAR row: portal URL, portal type, Status (Jobs found / No matching jobs / Blocked / Portal not found), kitni jobs, kya check hua |
| **Summary** | Kitni rows check hui vs input (MATCH hona chahiye), status count, High/Med/Low count, **Top 5 jobs** |

## Kaise kaam karta hai
1. Excel ki har row padhta hai (`Company`, `City`, `Website`, `Match for Aman`).
2. `portals.yaml` se us company ke portal ki recipe leta hai (Workday, Eightfold, Oracle, SmartRecruiters,
   Lever, SuccessFactors, Darwinbox, Keka/Zwayam/custom JSON, etc.). Sab recipes 13-Nov-2026 ko test ki gayi thi.
   Yaml mein na ho to website se careers page khud dhoondta hai.
3. Noida/Gurugram ki saari jobs (ya "embedded"/"firmware" search) nikalta hai, engineering jobs ka **JD kholta hai**.
4. Rakhta hai sirf wo jobs jinke title ya JD mein embedded / firmware / BSP / RTOS / AUTOSAR / CAN / UDS / MCU /
   device driver / bootloader / embedded Linux / HIL jaise kaam hon, aur location Delhi NCR ho.
   "2 Locations" / sirf "India" wali jobs **[Location unclear]** ke saath rakhi jaati hain.
5. JD se experience ("2-4 years"), skills (CAN, UDS, STM32, RTOS, BSP...) nikalta hai aur match score banata hai:
   aapki skills JD mein kitni hain + company ka domain (automotive > semiconductor > power/EV) + experience
   (3 saal tak = boost, 5+ saal = **[Senior – stretch]**).

## Match ka matlab
- **High**: aapki core skills (CAN/UDS/RTOS/MCU/BSP...) + experience fit. Pehle inhe apply karo.
- **Med**: embedded/firmware kaam hai, lekin domain alag ya experience thoda zyada.
- **Low**: testing/validation, hardware-leaning, ya senior role.
Ye keyword-based score hai. Apply karne se pehle JD zaroor padho.

## Status ka matlab (Portal Status sheet)
- **Blocked**: portal ne script ko mana kiya (403 / captcha / error). Notes mein link hai, browser mein kholo.
  Abhi known: NXP (Workday scripts ko 500/422 deta hai), Statcon (Cloudflare), GlobalLogic (Incapsula),
  Tech Mahindra (postback-only page), Panasonic (Akamai), Exicom (private API), Amber (sirf Naukri).
  Tessolve / Spark Minda (Darwinbox) kabhi kabhi rate-limit karte hain. Agle run mein aksar chal jaate hain.
- **Portal not found**: company ka koi job listing page nahi (sirf email / resume form). Notes mein email/link hai.
- Kuch companies (Yamaha, JBM, Lumax, Honda, Napino, Sona...) `manual` hain. Script unhe check nahi karta,
  sirf yaad dilata hai. Kabhi kabhi khud dekh lena.

## Jab kisi company ka portal badal jaaye
Portal Status mein "Blocked" ya error aaye, ya jobs aani band ho jaayein:
1. Company ka careers page browser mein kholo, job list wala page dhoondo.
2. `portals.yaml` mein us company ki entry update karo (file ke top pe har type ka format likha hai).
   Workday wale portal ke liye bas `url:` badalna kaafi hai.

## Limits (seedhi baat)
- LinkedIn / Naukri / Indeed scrape **nahi** karta (terms allow nahi karte). Unke liye job alerts lagao.
- Captcha / bot-block wale portals ko bypass nahi karta. Wo "Blocked" dikhte hain, unhe khud check karo.
- Kuch companies purani postings (2021-2024) bhi "active" dikhati hain. Posted date dekh ke decide karo.
- Match score keyword-based hai. Kabhi kabhi galat High/Low aa sakta hai.

Purana script: `job_checker_v1_backup.py` (backup).
