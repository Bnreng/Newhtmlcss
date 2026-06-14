import urllib.parse as up
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def li(q):  # LinkedIn, last 7 days (r604800), Saudi Arabia (captures Jeddah + remote-KSA)
    return ("https://www.linkedin.com/jobs/search/?keywords=" + up.quote(q) +
            "&location=" + up.quote("Saudi Arabia") + "&f_TPR=r604800")

def ind(q):  # Indeed Saudi Arabia, last 7 days
    return ("https://sa.indeed.com/jobs?q=" + up.quote(q) + "&l=" +
            up.quote("Saudi Arabia") + "&fromage=7")

# title, match, location, keyword1, keyword2, keyword3, variant
rows = [
 ("IT Systems Administrator",95,"Jeddah / Remote-KSA","Windows Server 2022","VMware vSphere","Active Directory","Systems & Infrastructure"),
 ("Systems Administrator",94,"Jeddah / Remote-KSA","Windows Server","Linux (Ubuntu/CentOS)","Backup & Disaster Recovery","Systems & Infrastructure"),
 ("IT Administrator",91,"Jeddah","Active Directory & GPO","RBAC","Vendor Management","Systems & Infrastructure"),
 ("Windows Systems Administrator",90,"Jeddah / Remote-KSA","Windows Server 2019/2022","Group Policy","Hyper-V/VMware","Systems & Infrastructure"),
 ("IT Specialist",89,"Jeddah","Infrastructure Support","Cisco Networking","End-user Support","Systems & Infrastructure"),
 ("IT Infrastructure Engineer",88,"Jeddah / Remote-KSA","VMware Virtualization","Server Infrastructure","Uptime Monitoring","Systems & Infrastructure"),
 ("IT Support Engineer",86,"Jeddah","Troubleshooting","First-Contact Resolution","Hardware/Software","IT Support"),
 ("IT Operations Engineer",84,"Jeddah / Remote-KSA","ITIL","Incident Management","Proactive Monitoring","Systems & Infrastructure"),
 ("Network Administrator",84,"Jeddah","Cisco Switches & Routers","VLANs / VPN","DNS / DHCP","Network Administration"),
 ("System Engineer",83,"Jeddah","Server Deployment","Active Directory","Virtualization","Systems & Infrastructure"),
 ("Desktop Support Engineer",83,"Jeddah","End-user Support","User Onboarding","Ticketing/SLA","IT Support"),
 ("Virtualization Administrator (VMware)",82,"Jeddah / Remote-KSA","VMware vSphere","Docker","VM Lifecycle","Systems & Infrastructure"),
 ("Network Support Engineer",80,"Jeddah","Firewalls","Enterprise Wi-Fi / AP","LAN/WAN","Network Administration"),
 ("Help Desk Team Lead",80,"Jeddah","SLA Management","Escalation","Knowledge Base","IT Support"),
 ("Network Engineer",79,"Jeddah","Cisco / TCP-IP","Routing & Switching","LAN/WAN","Network Administration"),
 ("Data Center Technician",78,"Jeddah","Server Hardware","Managed Switches","Uptime/Availability","Systems & Infrastructure"),
 ("SOC Analyst (Tier 1)",72,"Jeddah / Remote-KSA","Wazuh SIEM","Log Analysis","Incident Response","Cybersecurity / SOC"),
 ("Cybersecurity Analyst",71,"Jeddah / Remote-KSA","Threat Detection","Wazuh / Splunk","Security Hardening","Cybersecurity / SOC"),
 ("Information Security Analyst",70,"Jeddah / Remote-KSA","SIEM Operations","Incident Response","RBAC","Cybersecurity / SOC"),
 ("SIEM Analyst",69,"Jeddah / Remote-KSA","Wazuh","Splunk","Log Correlation","Cybersecurity / SOC"),
]
rows.sort(key=lambda r: r[1], reverse=True)

wb = Workbook()
ws = wb.active
ws.title = "Job Search Tracker"

H = ["Rank","Job Title","Company Name","Location (or Remote)","Match Score (1-100)",
     "3 Keywords to Emphasize","LinkedIn – Last 7 Days (click)","Indeed SA – Last 7 Days (click)",
     "Resume Variant","Listing URL (paste as you apply)","Status"]

hdr_fill = PatternFill("solid", fgColor="1F4E78")
hdr_font = Font(bold=True, color="FFFFFF", size=11)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin,right=thin,top=thin,bottom=thin)

ws.append(H)
for c in range(1,len(H)+1):
    cell = ws.cell(row=1,column=c)
    cell.fill = hdr_fill; cell.font = hdr_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

def score_color(s):
    if s>=90: return "C6EFCE"
    if s>=80: return "FFEB9C"
    if s>=70: return "FCE4D6"
    return "F2F2F2"

for i,(title,match,loc,k1,k2,k3,variant) in enumerate(rows, start=1):
    r = i+1
    ws.cell(row=r,column=1,value=i)
    ws.cell(row=r,column=2,value=title)
    ws.cell(row=r,column=3,value="(loads in live search →)")
    ws.cell(row=r,column=4,value=loc)
    sc = ws.cell(row=r,column=5,value=match)
    sc.fill = PatternFill("solid", fgColor=score_color(match))
    sc.alignment = Alignment(horizontal="center")
    ws.cell(row=r,column=6,value=f"{k1}; {k2}; {k3}")
    lc = ws.cell(row=r,column=7,value="Open LinkedIn (7d)")
    lc.hyperlink = li(title.split(" (")[0]); lc.font = Font(color="0563C1", underline="single")
    ic = ws.cell(row=r,column=8,value="Open Indeed (7d)")
    ic.hyperlink = ind(title.split(" (")[0]); ic.font = Font(color="0563C1", underline="single")
    ws.cell(row=r,column=9,value=variant)
    ws.cell(row=r,column=10,value="")
    ws.cell(row=r,column=11,value="To apply")
    for c in range(1,len(H)+1):
        ws.cell(row=r,column=c).border = border
        ws.cell(row=r,column=c).alignment = Alignment(vertical="center", wrap_text=(c in (6,)))

widths = [6,34,24,20,12,40,22,22,24,30,12]
for i,w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(H))}{len(rows)+1}"

# README sheet
rd = wb.create_sheet("README – How to use")
notes = [
 ("Bandar Alzahrani — Job Search Workbook", True),
 ("", False),
 ("HONEST DATA NOTE — please read:", True),
 ("LinkedIn, Indeed and Bayt block automated scraping (HTTP 403 / login wall), so this workbook does NOT", False),
 ("contain invented company names or fake listing URLs. Instead, each row gives you a LIVE search link that", False),
 ("is pre-filtered to your title + Saudi Arabia + posted in the last 7 days. Click it while logged in and", False),
 ("the real, current postings load in your browser. These links self-update, so they never go stale.", False),
 ("", False),
 ("HOW TO USE:", True),
 ("1. Sort tab 'Job Search Tracker' is already sorted by Match Score (highest first).", False),
 ("2. For each title, click 'Open LinkedIn (7d)' and 'Open Indeed (7d)' to see live last-7-day jobs.", False),
 ("3. As you find a real listing, paste its company + URL into the Company / Listing URL columns.", False),
 ("4. Use the matching 'Resume Variant' (see Resume_Variants.md) for that application.", False),
 ("5. Emphasize the 3 keywords shown for that title in your CV and cover note (ATS keyword match).", False),
 ("", False),
 ("MATCH SCORE LOGIC (1-100):", True),
 ("90-95 = near-exact fit to your current IT Systems Admin experience.", False),
 ("80-89 = strong fit; you cover the core, minor stretch on title/seniority.", False),
 ("70-79 = good fit with framing; leans on your certs/degree more than hands-on years.", False),
 ("69-72 = security/SOC roles: backed by your Wazuh SIEM build + Cybersecurity degree + ISC2/Ethical Hacking,", False),
 ("        but you have project experience rather than a dedicated SOC job title — be ready to speak to that.", False),
 ("", False),
 ("WANT TRUE PER-LISTING ROWS?", True),
 ("Paste the specific listings you want scored (copy the job text), and I'll fill exact company, location,", False),
 ("per-job match score, and a fully custom resume for each one.", False),
]
for i,(t,b) in enumerate(notes, start=1):
    c = rd.cell(row=i,column=1,value=t)
    if b: c.font = Font(bold=True, size=12 if i==1 else 11, color="1F4E78")
rd.column_dimensions["A"].width = 110

wb.save("/home/user/Newhtmlcss/job_search/Bandar_JobSearch_Tracker.xlsx")
print("xlsx written, rows:", len(rows))

# CSV fallback
import csv
with open("/home/user/Newhtmlcss/job_search/Bandar_JobSearch_Tracker.csv","w",newline="") as f:
    w = csv.writer(f)
    w.writerow(["Rank","Job Title","Company Name","Location (or Remote)","Match Score",
                "3 Keywords to Emphasize","LinkedIn Last-7-Days Link","Indeed Last-7-Days Link","Resume Variant"])
    for i,(title,match,loc,k1,k2,k3,variant) in enumerate(rows, start=1):
        w.writerow([i,title,"(loads in live search)",loc,match,f"{k1}; {k2}; {k3}",
                    li(title.split(' (')[0]), ind(title.split(' (')[0]), variant])
print("csv written")
